"""
Standalone Selenium helper for creating an external user category in eTeams Org. Structure.

Usage:
    from create_external_user_category import create_external_user_category
    result = create_external_user_category(driver)
    result = create_external_user_category(
        driver,
        category_name="xuyingtest外部用户分类001",
        remark="自动化测试备注",
    )

Precondition: ``driver`` is already logged in to eTeams. The helper will enter
``Org. Structure`` when needed, click the left-side ``外部用户分类设置`` menu,
click the top-right ``新建`` button, fill category information and save.
"""

from __future__ import annotations

import os
import secrets
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from create_new_external_user import (
    DEFAULT_IMPLICIT_WAIT_SECONDS,
    _body_text,
    _collect_visible_messages,
    _ensure_org_structure_page,
    _find_text_control_by_label,
    _safe_click,
    _set_text_control_value,
)

DEFAULT_TEST_EXTERNAL_USER_CATEGORY_PREFIX = os.environ.get(
    "ETEAMS_TEST_EXTERNAL_USER_CATEGORY_PREFIX",
    "xuyingtest外部用户分类",
)


@dataclass(frozen=True)
class ExternalUserCategoryData:
    """Generated values used to fill the new external-user-category form."""

    name: str
    remark: str


def _generate_external_user_category_data(
    category_name: str = "",
    remark: str = "",
) -> ExternalUserCategoryData:
    """Generate unique, recognizable values for one category creation run."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"
    name = (category_name or f"{DEFAULT_TEST_EXTERNAL_USER_CATEGORY_PREFIX}{compact_suffix}").strip()
    remark_value = (remark or f"自动化测试外部用户分类备注 {compact_suffix}").strip()
    return ExternalUserCategoryData(name=name, remark=remark_value)


def _is_external_user_category_setting_open(driver: webdriver.Chrome) -> bool:
    """Return True when 外部用户分类设置 is open and the top-right 新建 button is available."""
    try:
        return bool(
            driver.execute_script(
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
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                const bodyText = normalize(document.body.innerText || document.body.textContent);
                if (!bodyText.includes('外部用户分类设置')) return false;

                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!['新建', '新增', 'New', 'Add'].includes(text)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 500 || rect.top < 40 || rect.top > 180) continue;
                  return true;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _click_left_external_user_category_setting_menu(driver: webdriver.Chrome) -> str:
    """Click the left-side Org. Structure secondary menu item 外部用户分类设置."""
    try:
        return str(
            driver.execute_script(
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
                const fireClick = (el) => {
                  const rect = el.getBoundingClientRect();
                  const x = Math.floor(rect.left + rect.width / 2);
                  const y = Math.floor(rect.top + rect.height / 2);
                  for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    el.dispatchEvent(new MouseEvent(type, {
                      bubbles: true,
                      cancelable: true,
                      view: window,
                      clientX: x,
                      clientY: y
                    }));
                  }
                  if (typeof el.click === 'function') el.click();
                };
                const labels = [
                  '外部用户分类设置',
                  'External User Category Settings',
                  'External User Category',
                  'External User Categories'
                ];
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title'));
                  if (!labels.includes(text)) continue;
                  const clickable = raw.closest('.ui-menu-list-item, li, a, button, [role="menuitem"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  if (rect.left > 360 || rect.top < 50) continue;
                  candidates.push({
                    element: clickable,
                    text,
                    left: rect.left,
                    top: rect.top,
                    className: String(clickable.className || '')
                  });
                }
                candidates.sort((a, b) =>
                  a.left - b.left
                  || a.top - b.top
                  || (b.className.includes('ui-menu-list-item') ? 1 : 0)
                     - (a.className.includes('ui-menu-list-item') ? 1 : 0)
                );
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent) || candidates[0].text;
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_external_user_category_setting(driver: webdriver.Chrome) -> str:
    """Open 外部用户分类设置 from the left-side Org. Structure menu."""
    org_result = _ensure_org_structure_page(driver)
    if not org_result.startswith("已") and "已验证进入" not in org_result:
        return f"未能进入 Org. Structure，无法打开外部用户分类设置：{org_result}"

    if _is_external_user_category_setting_open(driver):
        return "已在左侧「外部用户分类设置」页面。"

    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: "外部用户分类设置" in _body_text(drv)
        )
    except TimeoutException:
        return "未能进入外部用户分类设置：Org. Structure 页面未出现左侧「外部用户分类设置」菜单。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_left_external_user_category_setting_menu(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                _is_external_user_category_setting_open
            )
            return f"已点击左侧二级菜单 {last_clicked}，并已验证进入外部用户分类设置页面。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能进入外部用户分类设置：已尝试点击左侧「外部用户分类设置」"
        f"{'（最后点击文本：' + last_clicked + '）' if last_clicked else '，但未找到可点击菜单项'}。"
    )


def _new_external_user_category_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the new external-user-category dialog/drawer is visible."""
    try:
        return bool(
            driver.execute_script(
                """
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
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                const selector = [
                  '.ui-dialog-wrap-right',
                  '.ui-dialog-wrap',
                  '.ui-modal',
                  '.modal',
                  '.ant-modal',
                  '.el-dialog',
                  '[role="dialog"]'
                ].join(',');
                for (const el of Array.from(document.querySelectorAll(selector))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if ((text.includes('新建') || text.includes('新增'))
                    && (text.includes('外部用户分类') || text.includes('分类名称') || text.includes('分类'))) {
                    return true;
                  }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _get_visible_new_external_user_category_dialog(driver: webdriver.Chrome):
    """Return the visible new-category dialog/drawer WebElement, if present."""
    try:
        return driver.execute_script(
            """
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
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const selector = [
              '.ui-dialog-wrap-right',
              '.ui-dialog-wrap',
              '.ui-modal',
              '.modal',
              '.ant-modal',
              '.el-dialog',
              '[role="dialog"]'
            ].join(',');
            const dialogs = Array.from(document.querySelectorAll(selector))
              .filter((el) => {
                if (!visible(el)) return false;
                const text = normalize(el.innerText || el.textContent);
                return (text.includes('新建') || text.includes('新增'))
                  && (text.includes('外部用户分类') || text.includes('分类名称') || text.includes('分类'));
              });
            dialogs.sort((a, b) =>
              b.getBoundingClientRect().left - a.getBoundingClientRect().left
              || a.getBoundingClientRect().top - b.getBoundingClientRect().top
            );
            return dialogs[0] || null;
            """
        )
    except Exception:
        return None


def _click_new_external_user_category_button(driver: webdriver.Chrome) -> str:
    """Click the top-right 新建 button on 外部用户分类设置."""
    try:
        return str(
            driver.execute_script(
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
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                const fireClick = (el) => {
                  const rect = el.getBoundingClientRect();
                  const x = Math.floor(rect.left + rect.width / 2);
                  const y = Math.floor(rect.top + rect.height / 2);
                  for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    el.dispatchEvent(new MouseEvent(type, {
                      bubbles: true,
                      cancelable: true,
                      view: window,
                      clientX: x,
                      clientY: y
                    }));
                  }
                  if (typeof el.click === 'function') el.click();
                };
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button.ui-btn, button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!['新建', '新增', 'New', 'Add'].includes(text)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 500 || rect.top < 40 || rect.top > 180) continue;
                  const disabled = el.disabled
                    || el.getAttribute('aria-disabled') === 'true'
                    || /disabled/.test(String(el.className || '').toLowerCase());
                  if (disabled) continue;
                  candidates.push({ element: el, text, left: rect.left, top: rect.top });
                }
                candidates.sort((a, b) => b.left - a.left || a.top - b.top);
                const target = candidates[0]?.element;
                if (!target) return '';
                fireClick(target);
                return candidates[0].text;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_new_external_user_category_dialog(driver: webdriver.Chrome) -> str:
    """Click 新建 and verify the category dialog opens."""
    if not _is_external_user_category_setting_open(driver):
        return "未在「外部用户分类设置」页面，无法点击「新建」。"

    if _new_external_user_category_dialog_is_open(driver):
        return "新建外部用户分类弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_external_user_category_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_external_user_category_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建外部用户分类弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建外部用户分类弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到右上角「新建」按钮。'}"
    )


def _find_first_required_text_control(driver: webdriver.Chrome):
    """Find the first required text-like input in the new category dialog."""
    dialog = _get_visible_new_external_user_category_dialog(driver)
    if dialog is None:
        return None
    try:
        return driver.execute_script(
            """
            const dialog = arguments[0];
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
            const labelCandidates = Array.from(dialog.querySelectorAll('.ui-formItem-label-required, .required, label, span, div'));
            const candidates = [];
            for (const label of labelCandidates) {
              if (!visible(label)) continue;
              const labelText = String(label.innerText || label.textContent || '').trim();
              const className = String(label.className || '').toLowerCase();
              if (!labelText.includes('*') && !labelText.includes('＊') && !className.includes('required')) continue;
              const labelRect = label.getBoundingClientRect();
              let container = label.closest('.ui-formItem') || label.closest('.ui-form-col') || label.parentElement;
              for (let depth = 0; container && container !== dialog && depth < 5; depth += 1, container = container.parentElement) {
                for (const input of Array.from(container.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                  if (!visible(input) || input.disabled || input.readOnly) continue;
                  const tag = String(input.tagName || '').toLowerCase();
                  const type = String(input.getAttribute('type') || '').toLowerCase();
                  if (tag === 'input' && ['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) continue;
                  const rect = input.getBoundingClientRect();
                  if (rect.left <= labelRect.left || Math.abs(rect.top - labelRect.top) > 55) continue;
                  candidates.push({ element: input, top: rect.top, left: rect.left });
                }
                if (candidates.length) break;
              }
            }
            candidates.sort((a, b) => a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            dialog,
        )
    except Exception:
        return None


def _find_external_user_category_name_input(driver: webdriver.Chrome):
    """Return the input immediately to the right of 新建分类「分类名称」."""
    dialog = _get_visible_new_external_user_category_dialog(driver)
    if dialog is None:
        return None

    try:
        return driver.execute_script(
            """
            const dialog = arguments[0];
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
            const cleanLabel = (value) => normalize(value)
              .replace(/[\\*＊:：]/g, '')
              .trim();
            const usableInput = (input) => {
              if (!visible(input) || input.disabled || input.readOnly) return false;
              const tag = String(input.tagName || '').toLowerCase();
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (tag === 'input' && ['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) return false;
              return tag === 'input' || tag === 'textarea' || input.isContentEditable;
            };

            const candidates = [];
            const labels = Array.from(dialog.querySelectorAll(
              '.ui-formItem-label-span, .ui-formItem-label, label, span, div'
            ));
            for (const label of labels) {
              if (!visible(label)) continue;
              const text = cleanLabel(
                label.innerText
                || label.textContent
                || label.getAttribute('title')
                || label.getAttribute('aria-label')
              );
              if (text !== '分类名称') continue;
              const labelRect = label.getBoundingClientRect();

              let container = label.closest('.ui-formItem')
                || label.closest('.ui-form-row')
                || label.closest('.ui-layout-row')
                || label.parentElement;
              for (let depth = 0; container && container !== dialog && depth < 7; depth += 1, container = container.parentElement) {
                for (const input of Array.from(container.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                  if (!usableInput(input)) continue;
                  const rect = input.getBoundingClientRect();
                  const sameRow = Math.abs((rect.top + rect.height / 2) - (labelRect.top + labelRect.height / 2)) <= 40;
                  if (!sameRow || rect.left <= labelRect.left) continue;
                  // Do not confuse it with 显示顺序/描述 fields lower down.
                  if (rect.top > labelRect.top + 45) continue;
                  candidates.push({
                    element: input,
                    score: 200 - Math.abs(rect.top - labelRect.top) - Math.abs(rect.left - labelRect.right) / 10,
                    top: rect.top,
                    left: rect.left
                  });
                }
                if (candidates.length) break;
              }
            }

            // Last-resort fallback based on the actual 新建分类 layout: the
            // first visible text input in the dialog is 分类名称; 显示顺序 already
            // has value 1 and 描述 is below it.
            if (!candidates.length) {
              for (const input of Array.from(dialog.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                if (!usableInput(input)) continue;
                const rect = input.getBoundingClientRect();
                const value = normalize(input.value || input.textContent || '');
                const placeholder = normalize(input.getAttribute('placeholder'));
                if (rect.left < 500 || rect.top < 260 || rect.top > 340) continue;
                if (value === '1') continue;
                candidates.push({
                  element: input,
                  score: 80 + (placeholder === '请输入' ? 20 : 0),
                  top: rect.top,
                  left: rect.left
                });
              }
            }

            candidates.sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            dialog,
        )
    except Exception:
        return None


def _fill_external_user_category_form(
    driver: webdriver.Chrome,
    data: ExternalUserCategoryData,
) -> str:
    """Fill category name and optional remark/description."""
    try:
        name_control = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_external_user_category_name_input(drv)
            or _find_text_control_by_label(
                drv,
                ["分类名称", "外部用户分类名称", "名称", "Name", "Category Name"],
                root_getter=_get_visible_new_external_user_category_dialog,
            )
            or _find_first_required_text_control(drv)
        )
    except TimeoutException:
        return "未找到新建外部用户分类弹窗中的分类名称/必填输入框。"

    if not _set_text_control_value(driver, name_control, data.name):
        return "填写「分类名称」失败。"

    filled = [f"分类名称={data.name}"]

    remark_control = _find_text_control_by_label(
        driver,
        ["备注信息", "备注", "描述", "说明", "Remark", "Remarks", "Description"],
        prefer_textarea=True,
        root_getter=_get_visible_new_external_user_category_dialog,
    )
    if remark_control is not None and _set_text_control_value(driver, remark_control, data.remark):
        filled.append(f"备注={data.remark}")
    elif remark_control is None:
        filled.append("备注字段未显示，已跳过")
    else:
        filled.append("备注字段填写失败，已继续")

    return "已填写新建外部用户分类表单：" + "；".join(filled) + "。"


def _click_save_new_external_user_category(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new external-user-category dialog."""
    dialog = _get_visible_new_external_user_category_dialog(driver)
    if dialog is None:
        return ""
    try:
        return str(
            driver.execute_script(
                """
                const dialog = arguments[0];
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
                const fireClick = (el) => {
                  const rect = el.getBoundingClientRect();
                  const x = Math.floor(rect.left + rect.width / 2);
                  const y = Math.floor(rect.top + rect.height / 2);
                  for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    el.dispatchEvent(new MouseEvent(type, {
                      bubbles: true,
                      cancelable: true,
                      view: window,
                      clientX: x,
                      clientY: y
                    }));
                  }
                  if (typeof el.click === 'function') el.click();
                };
                const candidates = [];
                for (const el of Array.from(dialog.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!['保存', 'Save', '确定', 'OK', 'Confirm'].includes(text)) continue;
                  const disabled = el.disabled
                    || el.getAttribute('aria-disabled') === 'true'
                    || /disabled/.test(String(el.className || '').toLowerCase());
                  if (disabled) continue;
                  const rect = el.getBoundingClientRect();
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.y - a.y || b.x - a.x);
                const target = candidates[0]?.element;
                if (!target) return '';
                fireClick(target);
                return candidates[0].text;
                """,
                dialog,
            )
        ).strip()
    except Exception:
        return ""


def _category_name_visible(driver: webdriver.Chrome, category_name: str) -> bool:
    """Return True if the category name appears outside visible dialogs."""
    try:
        return bool(
            driver.execute_script(
                """
                const categoryName = arguments[0];
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
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                for (const el of Array.from(document.body.querySelectorAll('td, .ui-table-cell, .ui-table-row, [role="row"], [role="cell"], a, span, div'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 220 || rect.top < 90) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (text && text.includes(categoryName) && text.length < 1000) return true;
                }
                return false;
                """,
                category_name,
            )
        )
    except Exception:
        return False


def _find_external_user_category_search_input(driver: webdriver.Chrome):
    """Return a search input on 外部用户分类设置, if one is available."""
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
              const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
              return dialog && visible(dialog);
            };
            const candidates = [];
            for (const input of Array.from(document.querySelectorAll('input, textarea'))) {
              if (!visible(input) || insideVisibleDialog(input) || input.disabled || input.readOnly) continue;
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (['hidden', 'password', 'checkbox', 'radio', 'file', 'button', 'submit', 'reset'].includes(type)) continue;
              const rect = input.getBoundingClientRect();
              if (rect.left < 500 || rect.top < 40 || rect.top > 220) continue;
              const attrs = normalize([
                input.getAttribute('placeholder'),
                input.getAttribute('title'),
                input.getAttribute('aria-label'),
                input.className,
                input.parentElement?.className,
                input.parentElement?.innerText
              ].join(' '));
              let score = rect.left / 10;
              if (/分类|名称|搜索|查询|请输入|search|name|category/i.test(attrs)) score += 200;
              if (/搜索|查询|search/i.test(attrs)) score += 120;
              if (/分类|名称|category|name/i.test(attrs)) score += 80;
              candidates.push({ element: input, score, left: rect.left, top: rect.top });
            }
            candidates.sort((a, b) => b.score - a.score || b.left - a.left || a.top - b.top);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _search_external_user_category_by_name(driver: webdriver.Chrome, category_name: str) -> bool:
    """Best-effort search/list verification for the saved category."""
    if _category_name_visible(driver, category_name):
        return True

    search_input = _find_external_user_category_search_input(driver)
    if search_input is None:
        return _category_name_visible(driver, category_name)

    try:
        _safe_click(driver, search_input)
        try:
            search_input.clear()
        except Exception:
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.BACKSPACE)
        if not _set_text_control_value(driver, search_input, category_name):
            search_input.send_keys(category_name)
        search_input.send_keys(Keys.ENTER)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const tag = String(input.tagName || '').toLowerCase();
            const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
            if (setter) setter.call(input, value); else input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            for (const type of ['keydown', 'keypress', 'keyup']) {
              input.dispatchEvent(new KeyboardEvent(type, {
                bubbles: true,
                cancelable: true,
                key: 'Enter',
                code: 'Enter',
                keyCode: 13,
                which: 13
              }));
            }
            """,
            search_input,
            category_name,
        )
        try:
            WebDriverWait(driver, 8, poll_frequency=0.4).until(
                lambda drv: _category_name_visible(drv, category_name)
            )
            return True
        except TimeoutException:
            return _category_name_visible(driver, category_name)
    except Exception:
        return _category_name_visible(driver, category_name)


def _save_external_user_category_and_wait(
    driver: webdriver.Chrome,
    category_name: str,
) -> str:
    """Save the category and wait for success/validation feedback."""
    clicked = _click_save_new_external_user_category(driver)
    if not clicked:
        return "未找到新建外部用户分类弹窗中的保存按钮。"

    duplicate_markers = ("已存在", "重复", "duplicate", "already exists", "Already exists")
    success_markers = ("保存成功", "新增成功", "创建成功", "操作成功", "success", "Success")
    validation_markers = ("必填", "不能为空", "请输入", "校验", "失败", "错误", "error", "Error")
    deadline = time.monotonic() + 15
    last_messages: list[str] = []
    dialog_closed_without_error = False

    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"外部用户分类 {category_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                return f"已保存外部用户分类 {category_name}，页面提示：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                return f"已点击 {clicked}，但页面提示校验/保存失败：{joined_messages}。"

        body_text = _body_text(driver)
        if any(marker in body_text for marker in duplicate_markers) and category_name in body_text:
            return f"外部用户分类 {category_name} 已存在，视为通过。"
        if category_name in body_text and not _new_external_user_category_dialog_is_open(driver):
            return f"已保存外部用户分类 {category_name}，页面列表中已可见。"
        if not _new_external_user_category_dialog_is_open(driver):
            dialog_closed_without_error = True
            break
        time.sleep(0.4)

    if _search_external_user_category_by_name(driver, category_name):
        return f"已保存外部用户分类 {category_name}，并已通过页面搜索/列表验证可见。"

    if dialog_closed_without_error:
        return (
            f"已点击 {clicked}，新建外部用户分类弹窗已关闭且未发现错误提示，"
            f"但未在页面上确认分类 {category_name} 可见。"
        )

    return (
        f"已点击 {clicked}，但未在页面上确认外部用户分类 {category_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def create_external_user_category(
    driver: webdriver.Chrome,
    category_name: str = "",
    remark: str = "",
) -> str:
    """
    Complete the Org. Structure -> 外部用户分类设置 -> 新建 flow.

    Args:
        driver: Existing Selenium Chrome driver that is already logged in.
        category_name: 分类名称；留空时自动生成 xuyingtest外部用户分类 + 时间戳。
        remark: 备注/描述；留空时自动生成自动化测试备注。
    """
    data = _generate_external_user_category_data(category_name=category_name, remark=remark)
    step_results: list[str] = []

    open_menu_result = _open_external_user_category_setting(driver)
    step_results.append(f"步骤2-外部用户分类设置：{open_menu_result}")
    if not open_menu_result.startswith("已"):
        return f"未能新建外部用户分类 {data.name}：" + " ".join(step_results)

    new_dialog_result = _open_new_external_user_category_dialog(driver)
    step_results.append(f"步骤3-新建：{new_dialog_result}")
    if not new_dialog_result.startswith(("已", "新建外部用户分类弹窗")):
        return f"未能新建外部用户分类 {data.name}：" + " ".join(step_results)

    fill_result = _fill_external_user_category_form(driver, data)
    step_results.append(f"步骤4-填写信息：{fill_result}")
    if not fill_result.startswith("已填写"):
        return f"未能新建外部用户分类 {data.name}：" + " ".join(step_results)

    save_result = _save_external_user_category_and_wait(driver, data.name)
    step_results.append(f"步骤5-保存：{save_result}")
    if save_result.startswith(("已保存", "外部用户分类")):
        return (
            f"已完成新建外部用户分类流程：分类名称={data.name}，备注={data.remark}。"
            + " ".join(step_results)
        )

    return f"未能确认新建外部用户分类 {data.name} 成功：" + " ".join(step_results)


__all__ = ["create_external_user_category"]
