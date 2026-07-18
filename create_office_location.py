"""
Standalone Selenium helper for creating a office location in eTeams Org. Structure.

Usage:
    from create_office_location import create_office_location
    result = create_office_location(driver)
    result = create_office_location(
        driver,
        location_name="xuyingtest办公地点001",
        remark="自动化测试备注",
    )

Precondition: ``driver`` is already logged in to eTeams. The helper will enter
``Org. Structure`` when needed, click the left-side ``办公地点`` menu, click the
top-right ``新建`` button, fill office location information and save.
"""

from __future__ import annotations

import os
import secrets
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from create_new_external_user import (
    _body_text,
    _collect_visible_messages,
    _ensure_org_structure_page,
    _find_text_control_by_label,
    _safe_click,
    _set_text_control_value,
)

DEFAULT_TEST_OFFICE_LOCATION_PREFIX = os.environ.get(
    "ETEAMS_TEST_OFFICE_LOCATION_PREFIX",
    "xuyingtest办公地点",
)


@dataclass(frozen=True)
class OfficeLocationData:
    """Generated values used to fill the new office-location form."""

    name: str
    remark: str


def _generate_office_location_data(
    location_name: str = "",
    remark: str = "",
) -> OfficeLocationData:
    """Generate unique, recognizable values for one office location run."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"
    name = (location_name or f"{DEFAULT_TEST_OFFICE_LOCATION_PREFIX}{compact_suffix}").strip()
    remark_value = (remark or f"自动化测试办公地点备注 {compact_suffix}").strip()
    return OfficeLocationData(name=name, remark=remark_value)


def _is_office_location_management_open(driver: webdriver.Chrome) -> bool:
    """Return True when 办公地点 is open and the top-right 新建 button is available."""
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
                if (!bodyText.includes('办公地点')) return false;

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


def _click_left_office_location_menu(driver: webdriver.Chrome) -> str:
    """Click the left-side Org. Structure secondary menu item 办公地点."""
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
                  '办公地点',
                  'Office Location Management',
                  'Office Locations',
                  'Office Location Management',
                  'Office Locations'
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


def _open_office_location_management(driver: webdriver.Chrome) -> str:
    """Open 办公地点 from the left-side Org. Structure menu."""
    org_result = _ensure_org_structure_page(driver)
    if not org_result.startswith("已") and "已验证进入" not in org_result:
        return f"未能进入 Org. Structure，无法打开办公地点：{org_result}"

    if _is_office_location_management_open(driver):
        return "已在左侧「办公地点」页面。"

    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: "办公地点" in _body_text(drv)
        )
    except TimeoutException:
        return "未能进入办公地点：Org. Structure 页面未出现左侧「办公地点」菜单。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_left_office_location_menu(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                _is_office_location_management_open
            )
            return f"已点击左侧二级菜单 {last_clicked}，并已验证进入办公地点页面。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能进入办公地点：已尝试点击左侧「办公地点」"
        f"{'（最后点击文本：' + last_clicked + '）' if last_clicked else '，但未找到可点击菜单项'}。"
    )


def _new_office_location_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the new department-group dialog/drawer is visible."""
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
                    && (text.includes('办公地点') || text.includes('地点名称') || text.includes('名称'))) {
                    return true;
                  }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _get_visible_new_office_location_dialog(driver: webdriver.Chrome):
    """Return the visible new office location dialog/drawer WebElement."""
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
                  && (text.includes('办公地点') || text.includes('地点名称') || text.includes('名称'));
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


def _click_new_office_location_button(driver: webdriver.Chrome) -> str:
    """Click the top-right 新建 button on 办公地点."""
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


def _open_new_office_location_dialog(driver: webdriver.Chrome) -> str:
    """Click 新建 and verify the office location dialog opens."""
    if not _is_office_location_management_open(driver):
        return "未在「办公地点」页面，无法点击「新建」。"

    if _new_office_location_dialog_is_open(driver):
        return "新建办公地点弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_office_location_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_office_location_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建办公地点弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建办公地点弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到右上角「新建」按钮。'}"
    )


def _find_office_location_name_input(driver: webdriver.Chrome):
    """Return the input immediately to the right of 新建办公地点「办公地点简称」."""
    dialog = _get_visible_new_office_location_dialog(driver)
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
            const labelMatches = (value) => {
              const text = cleanLabel(value);
              return ['办公地点简称', '办公地点名称', '办公地点名', '地点名称', '名称', 'Name', 'Location Name', 'Office Location Name']
                .some((label) => text === label || (text.includes(label) && text.length <= label.length + 8));
            };
            const usableInput = (input) => {
              if (!visible(input) || input.disabled || input.readOnly) return false;
              const tag = String(input.tagName || '').toLowerCase();
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (tag === 'input' && ['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) return false;
              return tag === 'input' || tag === 'textarea' || input.isContentEditable;
            };

            const candidates = [];
            const labelSelector = '.ui-formItem-label-span, .ui-formItem-label, label, span, div, [title], [aria-label]';
            for (const label of Array.from(dialog.querySelectorAll(labelSelector))) {
              if (!visible(label)) continue;
              const text = label.innerText || label.textContent || label.getAttribute('title') || label.getAttribute('aria-label');
              if (!labelMatches(text)) continue;
              const labelRect = label.getBoundingClientRect();

              let container = label.closest('.ui-formItem')
                || label.closest('.ui-form-row')
                || label.closest('.ui-layout-row')
                || label.parentElement;
              for (let depth = 0; container && container !== dialog && depth < 7; depth += 1, container = container.parentElement) {
                for (const input of Array.from(container.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                  if (!usableInput(input)) continue;
                  const rect = input.getBoundingClientRect();
                  const sameRow = Math.abs((rect.top + rect.height / 2) - (labelRect.top + labelRect.height / 2)) <= 45;
                  if (!sameRow || rect.left <= labelRect.left) continue;
                  candidates.push({
                    element: input,
                    score: 220 - depth * 5 - Math.abs(rect.top - labelRect.top) - Math.abs(rect.left - labelRect.right) / 10,
                    top: rect.top,
                    left: rect.left
                  });
                }
                if (candidates.length) break;
              }
            }

            // Last-resort fallback: choose the first empty text input near the top
            // of the dialog, avoiding numeric sort/order fields that often have value 1.
            if (!candidates.length) {
              for (const input of Array.from(dialog.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                if (!usableInput(input)) continue;
                const rect = input.getBoundingClientRect();
                const value = normalize(input.value || input.textContent || '');
                const cls = String(input.className || '').toLowerCase();
                if (rect.top < 120 || rect.top > 380) continue;
                if (value === '1' || /number/.test(cls)) continue;
                candidates.push({ element: input, score: 80, top: rect.top, left: rect.left });
              }
            }

            candidates.sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            dialog,
        )
    except Exception:
        return None


def _find_first_required_text_control(driver: webdriver.Chrome):
    """Find the first required text-like input in the new office location dialog."""
    dialog = _get_visible_new_office_location_dialog(driver)
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
            const candidates = [];
            for (const label of Array.from(dialog.querySelectorAll('.ui-formItem-label-required, .required, label, span, div'))) {
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
                  const value = String(input.value || input.textContent || '').trim();
                  const cls = String(input.className || '').toLowerCase();
                  if (value === '1' || /number/.test(cls)) continue;
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


def _find_office_location_city_input(driver: webdriver.Chrome):
    """Return the visible cascader input immediately to the right of「城市」."""
    dialog = _get_visible_new_office_location_dialog(driver)
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
              '.ui-formItem-label-span, .ui-formItem-label, label, span, div, [title], [aria-label]'
            ));
            for (const label of labels) {
              if (!visible(label)) continue;
              const text = cleanLabel(
                label.innerText
                || label.textContent
                || label.getAttribute('title')
                || label.getAttribute('aria-label')
              );
              if (text !== '城市') continue;
              const labelRect = label.getBoundingClientRect();
              let container = label.closest('.ui-formItem')
                || label.closest('.ui-form-row')
                || label.closest('.ui-layout-row')
                || label.parentElement;
              for (let depth = 0; container && container !== dialog && depth < 7; depth += 1, container = container.parentElement) {
                for (const input of Array.from(container.querySelectorAll('input.ui-cascader-input, input[placeholder="请选择"], input, textarea, [contenteditable="true"]'))) {
                  if (!usableInput(input)) continue;
                  const rect = input.getBoundingClientRect();
                  const sameRow = Math.abs((rect.top + rect.height / 2) - (labelRect.top + labelRect.height / 2)) <= 45;
                  if (!sameRow || rect.left <= labelRect.left) continue;
                  const className = String(input.className || '');
                  candidates.push({
                    element: input,
                    score: (className.includes('ui-cascader-input') ? 300 : 0)
                      + 220
                      - depth * 5
                      - Math.abs(rect.top - labelRect.top)
                      - Math.abs(rect.left - labelRect.right) / 10,
                    top: rect.top,
                    left: rect.left
                  });
                }
                if (candidates.length) break;
              }
            }

            if (!candidates.length) {
              for (const input of Array.from(dialog.querySelectorAll('input.ui-cascader-input, input[placeholder="请选择"]'))) {
                if (!usableInput(input)) continue;
                const rect = input.getBoundingClientRect();
                if (rect.top < 350 || rect.top > 560) continue;
                candidates.push({ element: input, score: 80, top: rect.top, left: rect.left });
              }
            }

            candidates.sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            dialog,
        )
    except Exception:
        return None


def _open_office_location_city_cascader(driver: webdriver.Chrome, city_input: Any) -> None:
    """Open the 城市 cascader dropdown."""
    try:
        _safe_click(driver, city_input)
    except Exception:
        pass
    try:
        driver.execute_script(
            """
            const input = arguments[0];
            input.scrollIntoView({ block: 'center', inline: 'center' });
            input.focus();
            const rect = input.getBoundingClientRect();
            const x = Math.floor(rect.left + rect.width / 2);
            const y = Math.floor(rect.top + rect.height / 2);
            for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
              input.dispatchEvent(new MouseEvent(type, {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: x,
                clientY: y
              }));
            }
            """,
            city_input,
        )
    except Exception:
        pass


def _click_office_location_cascader_option(
    driver: webdriver.Chrome,
    preferred_text: str,
    column_index: int,
    *,
    allow_first_visible: bool = False,
) -> str:
    """
    Click a visible cascader option in the requested column.

    Returns the clicked option text. If ``preferred_text`` is not visible and
    ``allow_first_visible`` is True, clicks the first enabled option in that
    cascader column as a fallback.
    """
    try:
        return str(
            driver.execute_script(
                """
                const preferredText = arguments[0];
                const columnIndex = Number(arguments[1]);
                const allowFirstVisible = Boolean(arguments[2]);
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
                const rawItems = Array.from(document.body.querySelectorAll(
                  'li.ui-cascader-menu-item, .ui-cascader-menu-item'
                ));
                const items = rawItems
                  .filter((el) => visible(el))
                  .map((el) => {
                    const rect = el.getBoundingClientRect();
                    const text = normalize(
                      el.getAttribute('title')
                      || el.innerText
                      || el.textContent
                      || el.getAttribute('aria-label')
                    );
                    const className = String(el.className || '').toLowerCase();
                    return { element: el, text, left: Math.round(rect.left), top: rect.top, className };
                  })
                  .filter((item) => item.text && !/disabled/.test(item.className));
                if (!items.length) return '';

                const columns = Array.from(new Set(items.map((item) => item.left)))
                  .sort((a, b) => a - b);
                const columnLeft = columns[columnIndex];
                if (columnLeft === undefined) return '';
                const columnItems = items
                  .filter((item) => Math.abs(item.left - columnLeft) <= 3)
                  .sort((a, b) => a.top - b.top);
                if (!columnItems.length) return '';

                let target = columnItems.find((item) => item.text === preferredText)?.element;
                let clickedText = preferredText;
                if (!target) {
                  const includesMatch = columnItems.find((item) =>
                    item.text.includes(preferredText)
                    || preferredText.includes(item.text)
                  );
                  if (includesMatch) {
                    target = includesMatch.element;
                    clickedText = includesMatch.text;
                  }
                }
                if (!target && allowFirstVisible) {
                  target = columnItems[0].element;
                  clickedText = columnItems[0].text;
                }
                if (!target) return '';
                fireClick(target);
                return clickedText;
                """,
                preferred_text,
                column_index,
                allow_first_visible,
            )
        ).strip()
    except Exception:
        return ""


def _visible_office_location_cascader_column_count(driver: webdriver.Chrome) -> int:
    """Return the number of currently visible cascader columns."""
    try:
        return int(
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
                const items = Array.from(document.body.querySelectorAll(
                  'li.ui-cascader-menu-item, .ui-cascader-menu-item'
                )).filter(visible);
                return new Set(items.map((el) => Math.round(el.getBoundingClientRect().left))).size;
                """
            )
            or 0
        )
    except Exception:
        return 0


def _get_office_location_city_display_value(
    driver: webdriver.Chrome,
    city_input: Any | None = None,
) -> str:
    """Return the displayed value of the 城市 cascader field."""
    try:
        if city_input is None:
            city_input = _find_office_location_city_input(driver)
        if city_input is None:
            return ""
        return str(
            driver.execute_script(
                """
                const input = arguments[0];
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

                const inputValue = normalize(input.value || input.getAttribute('value'));
                if (inputValue && inputValue !== '请选择') return inputValue;

                let container = input.closest('.ui-formItem')
                  || input.closest('.ui-form-col')
                  || input.closest('.ui-form-row')
                  || input.closest('.ui-layout-row')
                  || input.parentElement;
                for (let depth = 0; container && depth < 7; depth += 1, container = container.parentElement) {
                  for (const label of Array.from(container.querySelectorAll(
                    '.ui-cascader-picker-label, .ui-cascader-label, .ui-select-selection-selected-value, .ui-select-selection-rendered, [title]'
                  ))) {
                    if (!visible(label)) continue;
                    const text = normalize(
                      label.getAttribute('title')
                      || label.innerText
                      || label.textContent
                      || label.getAttribute('aria-label')
                    );
                    if (text && text !== '请选择' && text !== '城市' && !/^城市\\s*[:：]?$/.test(text)) {
                      return text;
                    }
                  }
                }

                container = input.closest('.ui-formItem')
                  || input.closest('.ui-form-col')
                  || input.closest('.ui-form-row')
                  || input.parentElement;
                const rowText = normalize(container?.innerText || container?.textContent || '');
                const stripped = normalize(rowText
                  .replace(/^城市\\s*[\\*＊]?\\s*[:：]?\\s*/, '')
                  .replace(/请选择/g, ''));
                return stripped
                  && stripped !== '城市'
                  && (stripped.includes('/') || stripped.includes('中国'))
                    ? stripped
                    : '';
                """,
                city_input,
            )
            or ""
        ).strip()
    except Exception:
        return ""


def _click_office_location_city_step(
    driver: webdriver.Chrome,
    city_input: Any,
    preferred_text: str,
    column_index: int,
    *,
    allow_first_visible: bool = False,
    timeout_seconds: float = 8.0,
) -> str:
    """Open the city cascader if needed and click one option in a column."""
    clicked = ""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline and not clicked:
        if _visible_office_location_cascader_column_count(driver) <= column_index:
            _open_office_location_city_cascader(driver, city_input)
            time.sleep(0.25)
        clicked = _click_office_location_cascader_option(
            driver,
            preferred_text,
            column_index,
            allow_first_visible=allow_first_visible,
        )
        if clicked:
            break
        time.sleep(0.3)
    return clicked


def _select_office_location_city_cascader(
    driver: webdriver.Chrome,
    *,
    country: str = "中国",
    city: str = "江苏省",
    district: str = "徐州市",
) -> str:
    """Select 城市 cascader: country -> province -> city/district/county."""
    try:
        city_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_office_location_city_input(drv)
        )
    except TimeoutException:
        return "未找到新建办公地点弹窗中的「城市」级联选择框。"

    selected: list[str] = []

    country_clicked = _click_office_location_city_step(
        driver,
        city_input,
        country,
        0,
        allow_first_visible=False,
    )
    if not country_clicked:
        return f"未能在城市级联选择框第 1 级选择「{country}」。"
    selected.append(country_clicked)
    time.sleep(0.5)

    city_clicked = _click_office_location_city_step(
        driver,
        city_input,
        city,
        1,
        allow_first_visible=True,
    )
    if not city_clicked:
        return f"未能在城市级联选择框第 2 级选择「{city}」。"
    selected.append(city_clicked)
    time.sleep(0.5)

    # Some tenants expose exactly 3 levels: 中国 -> 城市 -> 区县.
    # Others expose municipalities as 4 clicks: 中国 -> 省/直辖市 -> 城市 -> 区县.
    district_clicked = _click_office_location_city_step(
        driver,
        city_input,
        district,
        2,
        allow_first_visible=False,
        timeout_seconds=2.5,
    )
    if district_clicked:
        selected.append(district_clicked)
    else:
        city_drilldown_clicked = _click_office_location_city_step(
            driver,
            city_input,
            city,
            2,
            allow_first_visible=False,
            timeout_seconds=2.5,
        )
        if city_drilldown_clicked:
            selected.append(city_drilldown_clicked)
            time.sleep(0.5)
            district_clicked = _click_office_location_city_step(
                driver,
                city_input,
                district,
                3,
                allow_first_visible=True,
                timeout_seconds=3.0,
            )
            if not district_clicked:
                value = ""
                deadline = time.monotonic() + 3
                while time.monotonic() < deadline and not value:
                    value = _get_office_location_city_display_value(driver, city_input)
                    if value:
                        break
                    time.sleep(0.3)
                if value and (value.count("/") >= 2 or len(selected) >= 3):
                    return (
                        f"城市={'/'.join(selected)}"
                        f"（字段值：{value}；当前级联未出现独立区县列，已选择到可用末级）"
                    )
                return f"已选择 {'/'.join(selected)}，但未能继续选择区县「{district}」。"
            selected.append(district_clicked)
        else:
            district_clicked = _click_office_location_city_step(
                driver,
                city_input,
                district,
                2,
                allow_first_visible=True,
            )
            if not district_clicked:
                return f"未能在城市级联选择框第 3 级选择区县「{district}」。"
            selected.append(district_clicked)

    value = ""
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline and not value:
        value = _get_office_location_city_display_value(driver, city_input)
        if value:
            break
        time.sleep(0.3)

    if not value:
        return "已点击城市级联选项，但城市字段未显示已选值。"

    return f"城市={'/'.join(selected)}（字段值：{value}）"


def _fill_office_location_city_field(driver: webdriver.Chrome) -> str:
    """Fill the required 城市 cascader field."""
    return _select_office_location_city_cascader(driver)


def _fill_office_location_form(
    driver: webdriver.Chrome,
    data: OfficeLocationData,
) -> str:
    """Fill office location name, required city cascader, and optional remark/description."""
    try:
        name_control = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_office_location_name_input(drv)
            or _find_text_control_by_label(
                drv,
                [
                    "办公地点名称",
                    "办公地点简称",
                    "办公地点名",
                    "地点名称",
                    "名称",
                    "Name",
                    "Location Name",
                    "Office Location Name",
                ],
                root_getter=_get_visible_new_office_location_dialog,
            )
            or _find_first_required_text_control(drv)
        )
    except TimeoutException:
        return "未找到新建办公地点弹窗中的办公地点名称/必填输入框。"

    if not _set_text_control_value(driver, name_control, data.name):
        return "填写「办公地点简称」失败。"

    filled = [f"办公地点简称={data.name}"]

    full_name_control = _find_text_control_by_label(
        driver,
        ["办公地点全称", "地点全称", "全称", "Full Name", "Office Location Full Name"],
        root_getter=_get_visible_new_office_location_dialog,
    )
    if full_name_control is not None and _set_text_control_value(driver, full_name_control, data.name):
        filled.append(f"办公地点全称={data.name}")
    elif full_name_control is None:
        filled.append("办公地点全称字段未显示，已跳过")
    else:
        filled.append("办公地点全称填写失败，已继续")

    city_result = _fill_office_location_city_field(driver)
    if not city_result.startswith("城市="):
        return "填写新建办公地点表单失败：" + city_result
    filled.append(city_result)

    remark_control = _find_text_control_by_label(
        driver,
        ["备注信息", "备注", "描述", "说明", "Remark", "Remarks", "Description"],
        prefer_textarea=True,
        root_getter=_get_visible_new_office_location_dialog,
    )
    if remark_control is not None and _set_text_control_value(driver, remark_control, data.remark):
        filled.append(f"备注={data.remark}")
    elif remark_control is None:
        filled.append("备注字段未显示，已跳过")
    else:
        filled.append("备注字段填写失败，已继续")

    return "已填写新建办公地点表单：" + "；".join(filled) + "。"


def _click_save_new_office_location(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new office location dialog."""
    dialog = _get_visible_new_office_location_dialog(driver)
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


def _office_location_name_visible(driver: webdriver.Chrome, location_name: str) -> bool:
    """Return True if the department group name appears outside visible dialogs."""
    try:
        return bool(
            driver.execute_script(
                """
                const locationName = arguments[0];
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
                  if (text && text.includes(locationName) && text.length < 1000) return true;
                }
                return false;
                """,
                location_name,
            )
        )
    except Exception:
        return False


def _find_office_location_search_input(driver: webdriver.Chrome):
    """Return a search input on 办公地点, if one is available."""
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
              if (/办公地点|地点名称|名称|搜索|查询|请输入|search|name|location|office/i.test(attrs)) score += 200;
              if (/搜索|查询|search/i.test(attrs)) score += 120;
              if (/办公地点|地点名称|名称|location|office|name/i.test(attrs)) score += 80;
              candidates.push({ element: input, score, left: rect.left, top: rect.top });
            }
            candidates.sort((a, b) => b.score - a.score || b.left - a.left || a.top - b.top);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _search_office_location_by_name(driver: webdriver.Chrome, location_name: str) -> bool:
    """Best-effort search/list verification for the saved office location."""
    if _office_location_name_visible(driver, location_name):
        return True

    search_input = _find_office_location_search_input(driver)
    if search_input is None:
        return _office_location_name_visible(driver, location_name)

    try:
        _safe_click(driver, search_input)
        try:
            search_input.clear()
        except Exception:
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.BACKSPACE)
        if not _set_text_control_value(driver, search_input, location_name):
            search_input.send_keys(location_name)
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
            location_name,
        )
        try:
            WebDriverWait(driver, 8, poll_frequency=0.4).until(
                lambda drv: _office_location_name_visible(drv, location_name)
            )
            return True
        except TimeoutException:
            return _office_location_name_visible(driver, location_name)
    except Exception:
        return _office_location_name_visible(driver, location_name)


def _save_office_location_and_wait(
    driver: webdriver.Chrome,
    location_name: str,
) -> str:
    """Save the office location and wait for success/validation feedback."""
    clicked = ""
    click_deadline = time.monotonic() + 8
    while time.monotonic() < click_deadline:
        clicked = _click_save_new_office_location(driver)
        if clicked:
            break
        time.sleep(0.3)
    if not clicked:
        return "未找到新建办公地点弹窗中的保存按钮。"

    duplicate_markers = ("已存在", "重复", "duplicate", "already exists", "Already exists")
    success_markers = ("保存成功", "新增成功", "创建成功", "操作成功", "success", "Success")
    validation_markers = ("必填", "不能为空", "请输入", "请选择", "校验", "失败", "错误", "error", "Error")
    deadline = time.monotonic() + 15
    last_messages: list[str] = []
    dialog_closed_without_error = False

    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"办公地点 {location_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                return f"已保存办公地点 {location_name}，页面提示：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                return f"已点击 {clicked}，但页面提示校验/保存失败：{joined_messages}。"

        body_text = _body_text(driver)
        if any(marker in body_text for marker in duplicate_markers) and location_name in body_text:
            return f"办公地点 {location_name} 已存在，视为通过。"
        if location_name in body_text and not _new_office_location_dialog_is_open(driver):
            return f"已保存办公地点 {location_name}，页面列表中已可见。"
        if not _new_office_location_dialog_is_open(driver):
            dialog_closed_without_error = True
            break
        time.sleep(0.4)

    if _search_office_location_by_name(driver, location_name):
        return f"已保存办公地点 {location_name}，并已通过页面搜索/列表验证可见。"

    if dialog_closed_without_error:
        return (
            f"已点击 {clicked}，新建办公地点弹窗已关闭且未发现错误提示，"
            f"但未在页面上确认办公地点 {location_name} 可见。"
        )

    return (
        f"已点击 {clicked}，但未在页面上确认办公地点 {location_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def create_office_location(
    driver: webdriver.Chrome,
    location_name: str = "",
    remark: str = "",
) -> str:
    """
    Complete the Org. Structure -> 办公地点 -> 新建办公地点 flow.

    Args:
        driver: Existing Selenium Chrome driver that is already logged in.
        location_name: 办公地点名称；留空时自动生成 xuyingtest办公地点 + 时间戳。
        remark: 备注/描述；留空时自动生成自动化测试备注。
    """
    data = _generate_office_location_data(location_name=location_name, remark=remark)
    step_results: list[str] = []

    open_menu_result = _open_office_location_management(driver)
    step_results.append(f"步骤2-办公地点：{open_menu_result}")
    if not open_menu_result.startswith("已"):
        return f"未能新建办公地点 {data.name}：" + " ".join(step_results)

    new_dialog_result = _open_new_office_location_dialog(driver)
    step_results.append(f"步骤3-新建：{new_dialog_result}")
    if not new_dialog_result.startswith(("已", "新建办公地点弹窗")):
        return f"未能新建办公地点 {data.name}：" + " ".join(step_results)

    fill_result = _fill_office_location_form(driver, data)
    step_results.append(f"步骤4-填写信息：{fill_result}")
    if not fill_result.startswith("已填写"):
        return f"未能新建办公地点 {data.name}：" + " ".join(step_results)

    save_result = _save_office_location_and_wait(driver, data.name)
    step_results.append(f"步骤5-保存：{save_result}")
    if save_result.startswith(("已保存", "办公地点")):
        return (
            f"已完成新建办公地点流程：办公地点名称={data.name}，备注={data.remark}。"
            + " ".join(step_results)
        )

    return f"未能确认新建办公地点 {data.name} 成功：" + " ".join(step_results)


__all__ = ["create_office_location"]
