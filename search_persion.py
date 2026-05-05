"""
Standalone search helper for verifying a newly-created person in eTeams.

The filename intentionally follows the requested spelling: ``search_persion.py``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

if TYPE_CHECKING:
    from create_new_person import PersonData


def _get_human_resources_name_search_input(driver: webdriver.Chrome):
    """
    Return the real 人员姓名搜索框 on the 人力资源 page.

    The target field is the visible input whose placeholder contains
    ``请输入姓名``.  Keep this lookup stricter than the generic helper in
    ``create_new_person.py`` so we do not accidentally type into another
    toolbar input.
    """
    try:
        return driver.execute_script(
            """
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const visible = (el) => {
              if (!el) return false;
              const style = window.getComputedStyle(el);
              const rect = el.getBoundingClientRect();
              return style.display !== 'none'
                && style.visibility !== 'hidden'
                && Number(style.opacity || 1) > 0
                && rect.width > 0
                && rect.height > 0
                && rect.right > 0
                && rect.bottom > 0
                && rect.left < window.innerWidth
                && rect.top < window.innerHeight;
            };
            const insideVisibleDialog = (el) => {
              const dialog = el.closest(
                '.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, '
                + '.ant-modal, .el-dialog, [role="dialog"]'
              );
              return dialog && visible(dialog);
            };

            const newPersonLabels = [
              '新建人员',
              '新增人员',
              '新建用户',
              '新增用户',
              'New Person',
              'Add Person',
              'New Employee',
              'Add Employee',
              'New User',
              'Add User',
              'New Staff',
              'Add Staff'
            ];
            const newPersonButtons = [];
            for (const raw of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
              if (!visible(raw) || insideVisibleDialog(raw)) continue;
              const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title') || raw.getAttribute('aria-label'));
              if (!newPersonLabels.includes(text)) continue;
              const rect = raw.getBoundingClientRect();
              if (rect.left < 180 || rect.top < 40) continue;
              newPersonButtons.push({ rect });
            }
            newPersonButtons.sort((a, b) => b.rect.left - a.rect.left || a.rect.top - b.rect.top);
            const buttonRect = newPersonButtons[0]?.rect || null;

            const candidates = [];
            for (const input of Array.from(document.body.querySelectorAll('input, textarea'))) {
              if (!visible(input) || insideVisibleDialog(input)) continue;
              if (input.disabled || input.readOnly) continue;
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (['hidden', 'password', 'checkbox', 'radio', 'file', 'button', 'submit', 'reset'].includes(type)) continue;

              const placeholder = normalize(input.getAttribute('placeholder'));
              if (!/请输入\\s*姓名/.test(placeholder)) continue;

              const rect = input.getBoundingClientRect();
              if (rect.left < 160 || rect.top < 40) continue;

              let score = 1000;
              if (placeholder === '请输入姓名') score += 300;
              else if (placeholder.startsWith('请输入姓名')) score += 180;
              if (buttonRect) {
                const inputCenterY = rect.top + rect.height / 2;
                const buttonCenterY = buttonRect.top + buttonRect.height / 2;
                const sameRow = Math.abs(inputCenterY - buttonCenterY) <= 80;
                if (sameRow) score += 300;
                if (sameRow && rect.left >= buttonRect.right - 20) score += 200;
                score -= Math.abs(rect.left - buttonRect.right) / 5;
              }

              candidates.push({
                element: input,
                score,
                top: rect.top,
                left: rect.left,
                placeholder
              });
            }

            candidates.sort((a, b) =>
              b.score - a.score
              || a.top - b.top
              || a.left - b.left
            );
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _search_created_person_in_human_resources(
    driver: webdriver.Chrome,
    person: PersonData,
) -> str:
    """Search the 人力资源 page for the newly-created person and verify it appears."""
    from create_new_person import (
        _click_human_resources_search_trigger,
        _collect_visible_messages,
        _describe_human_resources_search_input,
        _ensure_org_structure_page,
        _fill_human_resources_search_input,
        _human_resources_search_result_snippet,
        _human_resources_tab_is_active,
        _new_person_dialog_is_open,
        _open_human_resources_tab,
        _open_org_maintenance,
        _person_detail_page_is_open,
    )

    if not _human_resources_tab_is_active(driver):
        org_result = _ensure_org_structure_page(driver)
        if not org_result.startswith(("已", "步骤", "选择")) and "已验证进入" not in org_result:
            return f"未在人力资源标签页，且未能进入 Org. Structure，无法搜索新建人员：{org_result}"

        org_maintenance_result = _open_org_maintenance(driver)
        if not org_maintenance_result.startswith("已"):
            return (
                "未在人力资源标签页，且未能进入「组织维护」，无法搜索新建人员："
                f"{org_maintenance_result}"
            )

        hr_result = _open_human_resources_tab(driver)
        if not hr_result.startswith("已"):
            return (
                "未在人力资源标签页，且未能进入「人力资源」标签页，无法搜索新建人员："
                f"{hr_result}"
            )

    try:
        WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: not _new_person_dialog_is_open(drv)
            and not _person_detail_page_is_open(drv, person.name)
        )
    except TimeoutException:
        return "新建人员弹窗或详情页仍未关闭，无法在人员列表中搜索。"

    try:
        search_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _get_human_resources_name_search_input(drv)
        )
    except TimeoutException:
        return (
            "未找到人力资源页面中 placeholder 包含「请输入姓名」"
            "且位于「新建人员」按钮右侧的姓名搜索框。"
        )

    input_label = _describe_human_resources_search_input(driver, search_input)
    input_value = _fill_human_resources_search_input(driver, search_input, person.name)
    if person.name not in input_value:
        return (
            "未能确认已在搜索框中输入新建人员姓名"
            f"（当前搜索框值：{input_value or '空'}）。"
        )

    clicked_search = _click_human_resources_search_trigger(driver, search_input)
    trigger_text = f"并点击 {clicked_search}" if clicked_search else "并按 Enter"

    try:
        snippet = WebDriverWait(driver, 15, poll_frequency=0.4).until(
            lambda drv: _human_resources_search_result_snippet(drv, person.name)
        )
    except TimeoutException:
        messages = _collect_visible_messages(driver, fast=True)
        return (
            f"已在搜索框输入 {person.name} {trigger_text}，"
            "但未在搜索结果中验证到该人员"
            f"{'；页面提示：' + '；'.join(messages) if messages else ''}。"
        )

    return (
        f"已在搜索框{f'（{input_label}）' if input_label else ''}输入 {person.name}，"
        f"{trigger_text}，并已在搜索结果中验证该人员可见"
        f"（结果片段：{snippet}）。"
    )


__all__ = ["_search_created_person_in_human_resources"]
