"""
Standalone Selenium helper for creating a new external enterprise in eTeams Org. Structure.

Usage:
    from create_new_external_user import create_new_enterprise
    result = create_new_enterprise(driver)  # auto-generates enterprise data
    result = create_new_enterprise(
        driver,
        enterprise_name="xuyingtest企业001",
        enterprise_full_name="xuyingtest企业001全称",
        remark="自动化测试备注",
    )

Precondition: ``driver`` is already logged in to eTeams. The helper will enter
``Org. Structure`` when needed, click the left-side ``外部组织维护`` menu, click
``新建企业``, fill ``企业名称``、``企业全称``、``备注信息`` and save.
"""

from __future__ import annotations

import os
import secrets
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from select_org_structure import select_org_structure as _select_org_structure

ORG_STRUCTURE_PATH = "/hrm/orgsetting/departmentSetting"
DEFAULT_IMPLICIT_WAIT_SECONDS = 5
DEFAULT_TEST_ENTERPRISE_PREFIX = os.environ.get(
    "ETEAMS_TEST_ENTERPRISE_PREFIX",
    "xuyingtest企业",
)
DEFAULT_TEST_EXTERNAL_DEPARTMENT_PREFIX = os.environ.get(
    "ETEAMS_TEST_EXTERNAL_DEPARTMENT_PREFIX",
    "xuyingtest部门",
)
DEFAULT_TEST_EXTERNAL_CONTACT_PREFIX = os.environ.get(
    "ETEAMS_TEST_EXTERNAL_CONTACT_PREFIX",
    "xuyingtest外部联系人",
)
TEST_ENTERPRISE_NAME_OVERRIDE = os.environ.get(
    "ETEAMS_TEST_ENTERPRISE_NAME",
    "",
).strip()
TEST_ENTERPRISE_FULL_NAME_OVERRIDE = os.environ.get(
    "ETEAMS_TEST_ENTERPRISE_FULL_NAME",
    "",
).strip()
TEST_ENTERPRISE_REMARK_OVERRIDE = os.environ.get(
    "ETEAMS_TEST_ENTERPRISE_REMARK",
    "",
).strip()


@dataclass(frozen=True)
class EnterpriseData:
    """Generated values used to fill the new-enterprise form."""

    name: str
    full_name: str
    remark: str


@dataclass(frozen=True)
class ExternalDepartmentData:
    """Generated values used to fill a department under an external enterprise."""

    name: str
    full_name: str
    remark: str


@dataclass(frozen=True)
class ExternalContactData:
    """Generated values used to fill a contact under an external department."""

    name: str
    mobile: str
    email: str
    phone: str
    remark: str


def _generate_enterprise_data(
    enterprise_name: str = "",
    enterprise_full_name: str = "",
    remark: str = "",
) -> EnterpriseData:
    """Generate unique, recognizable values for one enterprise creation run."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"

    name = (
        enterprise_name
        or TEST_ENTERPRISE_NAME_OVERRIDE
        or f"{DEFAULT_TEST_ENTERPRISE_PREFIX}{compact_suffix}"
    ).strip()
    full_name = (
        enterprise_full_name
        or TEST_ENTERPRISE_FULL_NAME_OVERRIDE
        or f"{name}全称"
    ).strip()
    remark_value = (
        remark
        or TEST_ENTERPRISE_REMARK_OVERRIDE
        or f"自动化测试备注 {compact_suffix}"
    ).strip()
    return EnterpriseData(name=name, full_name=full_name, remark=remark_value)


def _generate_external_department_data(
    department_name: str = "",
    department_full_name: str = "",
    department_remark: str = "",
) -> ExternalDepartmentData:
    """Generate unique, recognizable values for one external department run."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"
    default_name = f"{DEFAULT_TEST_EXTERNAL_DEPARTMENT_PREFIX}{compact_suffix}"

    name = (department_name or default_name).strip()
    full_name = (department_full_name or f"{name}全称").strip()
    remark_value = (department_remark or f"自动化测试部门备注 {compact_suffix}").strip()
    return ExternalDepartmentData(name=name, full_name=full_name, remark=remark_value)


def _generate_external_contact_data(
    contact_name: str = "",
    contact_mobile: str = "",
    contact_email: str = "",
    contact_phone: str = "",
    contact_remark: str = "",
) -> ExternalContactData:
    """Generate unique, recognizable values for one external contact run."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"
    safe_account = f"external{compact_suffix}"
    mobile_tail = f"{int(time.time() * 1000) % 10**9:09d}"

    name = (
        contact_name
        or f"{DEFAULT_TEST_EXTERNAL_CONTACT_PREFIX}{compact_suffix}"
    ).strip()
    mobile = (contact_mobile or f"13{mobile_tail}").strip()
    email = (contact_email or f"{safe_account}@example.com").strip()
    phone = (contact_phone or f"021{int(time.time() * 1000) % 10**8:08d}").strip()
    remark = (contact_remark or f"自动化测试外部联系人备注 {compact_suffix}").strip()
    return ExternalContactData(
        name=name,
        mobile=mobile,
        email=email,
        phone=phone,
        remark=remark,
    )


def _normalize_text(value: str | None) -> str:
    return " ".join(str(value or "").split())


def _body_text(driver: webdriver.Chrome) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def _safe_click(driver: webdriver.Chrome, element: Any) -> None:
    """Click an element, falling back to ActionChains/JavaScript if needed."""
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element,
    )
    try:
        element.click()
        return
    except (ElementClickInterceptedException, StaleElementReferenceException):
        pass
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(element).click().perform()
        return
    except Exception:
        pass

    driver.execute_script(
        """
        const el = arguments[0];
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
        """,
        element,
    )


def _collect_visible_messages(
    driver: webdriver.Chrome,
    *,
    fast: bool = False,
) -> list[str]:
    """Collect short visible status/error messages from common toast components."""
    if fast:
        driver.implicitly_wait(0)

    selectors = [
        ".ui-message",
        ".ui-toast",
        ".ui-notification",
        ".ant-message",
        ".ant-notification",
        ".el-message",
        "[role='alert']",
    ]
    messages: list[str] = []
    try:
        for selector in selectors:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                try:
                    text = element.text.strip()
                    if text and element.is_displayed() and text not in messages:
                        messages.append(text)
                except StaleElementReferenceException:
                    continue
    finally:
        if fast:
            driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)
    return messages


def _dismiss_transient_overlays(driver: webdriver.Chrome) -> None:
    """Close dropdowns or transient menus that may cover the target page."""
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.2)
    except Exception:
        pass


def _org_structure_page_is_open(driver: webdriver.Chrome) -> bool:
    """Return True when the current page is the Org. Structure page."""
    try:
        if ORG_STRUCTURE_PATH in driver.current_url:
            return True

        body_text = _body_text(driver)
        return (
            "组织架构设置" in body_text
            and "组织维护" in body_text
            and any(
                marker in body_text
                for marker in (
                    "外部组织维护",
                    "External Organization",
                    "Teams",
                    "Groups",
                    "Division Info",
                    "部门名称",
                    "部门全称",
                    "行政组织",
                    "人力资源",
                )
            )
        )
    except Exception:
        return False


def _ensure_org_structure_page(driver: webdriver.Chrome) -> str:
    """Enter Org. Structure through the real top-right eTeams menu if needed."""
    if _org_structure_page_is_open(driver):
        _dismiss_transient_overlays(driver)
        return "已在 Org. Structure 页面，并已验证页面打开。"

    select_result = _select_org_structure(driver)
    try:
        WebDriverWait(driver, 15, poll_frequency=0.4).until(_org_structure_page_is_open)
        _dismiss_transient_overlays(driver)
        return f"{select_result} 已验证进入 Org. Structure 页面。"
    except TimeoutException:
        return f"未能进入 Org. Structure 页面：{select_result}"


def _is_external_org_maintenance_open(driver: webdriver.Chrome) -> bool:
    """Return True when 外部组织维护 is selected and 新建企业 is actionable."""
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
                const textIsNewEnterprise = (text) =>
                  text === '新建企业'
                  || text === '新增企业'
                  || text === 'New Enterprise'
                  || text === 'Add Enterprise'
                  || text.includes('新建企业')
                  || text.includes('新增企业');
                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!textIsNewEnterprise(text)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 180 || rect.top < 40) continue;
                  return true;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _click_left_external_org_maintenance_menu(driver: webdriver.Chrome) -> str:
    """Click the left-side Org. Structure secondary menu item 外部组织维护."""
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
                  '外部组织维护',
                  '外部组织',
                  'External Organization Maintenance',
                  'External Org Maintenance',
                  'External Organization'
                ];
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title'));
                  if (!labels.includes(text)) continue;
                  const clickable = raw.closest('.ui-menu-list-item, li, a, button, [role="menuitem"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  if (rect.left > 320 || rect.top < 60) continue;
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


def _open_external_org_maintenance(driver: webdriver.Chrome) -> str:
    """Step 3: open 外部组织维护 from the left-side Org. Structure menu."""
    if not _org_structure_page_is_open(driver):
        return "未在 Org. Structure 页面，无法进入左侧「外部组织维护」。"

    _dismiss_transient_overlays(driver)

    if _is_external_org_maintenance_open(driver):
        return "已在左侧「外部组织维护」页面。"

    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: "外部组织维护" in _body_text(drv)
            or "External Organization" in _body_text(drv)
        )
    except TimeoutException:
        return "未能进入外部组织维护：Org. Structure 页面未出现左侧「外部组织维护」菜单。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_left_external_org_maintenance_menu(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                _is_external_org_maintenance_open
            )
            return f"已点击左侧二级菜单 {last_clicked}，并已验证进入外部组织维护页面。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能进入外部组织维护：已尝试点击左侧「外部组织维护」"
        f"{'（最后点击文本：' + last_clicked + '）' if last_clicked else '，但未找到可点击菜单项'}。"
    )


def _new_enterprise_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the 新建企业 dialog/drawer is visible."""
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
                const dialogSelector = [
                  '.ui-dialog-wrap-right',
                  '.ui-dialog-wrap',
                  '.ui-modal',
                  '.modal',
                  '.ant-modal',
                  '.el-dialog',
                  '[role="dialog"]'
                ].join(',');
                const titleMarkers = [
                  '新建企业',
                  '新增企业',
                  'New Enterprise',
                  'Add Enterprise'
                ];
                const fieldMarkers = ['企业名称', '企业全称', 'Enterprise Name'];
                for (const el of Array.from(document.querySelectorAll(dialogSelector))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if (titleMarkers.some((marker) => text.includes(marker))
                    && fieldMarkers.some((marker) => text.includes(marker))) {
                    return true;
                  }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _get_visible_new_enterprise_dialog(driver: webdriver.Chrome):
    """Return the visible 新建企业 dialog/drawer WebElement, if present."""
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
            const titleMarkers = ['新建企业', '新增企业', 'New Enterprise', 'Add Enterprise'];
            const fieldMarkers = ['企业名称', '企业全称', 'Enterprise Name'];
            const dialogs = Array.from(document.querySelectorAll(selector))
              .filter((el) => {
                if (!visible(el)) return false;
                const text = normalize(el.innerText || el.textContent);
                return titleMarkers.some((marker) => text.includes(marker))
                  && fieldMarkers.some((marker) => text.includes(marker));
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


def _click_new_enterprise_button(driver: webdriver.Chrome) -> str:
    """Click the page top-right 新建企业 button."""
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
                const insideVisibleDialog = (el) => {
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                const textIsNewEnterprise = (text) =>
                  text === '新建企业'
                  || text === '新增企业'
                  || text === 'New Enterprise'
                  || text === 'Add Enterprise'
                  || text.includes('新建企业')
                  || text.includes('新增企业');
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!textIsNewEnterprise(text)) continue;
                  const rect = el.getBoundingClientRect();
                  // The requested button is on the page toolbar/top-right area.
                  if (rect.left < 180 || rect.top < 40 || rect.top > 320) continue;
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.x - a.x || a.y - b.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_new_enterprise_dialog(driver: webdriver.Chrome) -> str:
    """Step 4: click 新建企业 and verify the dialog/drawer opens."""
    if not _is_external_org_maintenance_open(driver):
        return "未在外部组织维护页面，无法点击「新建企业」。"

    if _new_enterprise_dialog_is_open(driver):
        return "新建企业弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_enterprise_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_enterprise_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建企业弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建企业弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到页面右上角可点击的「新建企业」按钮。'}"
    )


def _find_text_control_by_label(
    driver: webdriver.Chrome,
    labels: list[str],
    *,
    prefer_textarea: bool = False,
    root_getter=None,
):
    """Find an input/textarea/contenteditable control by a nearby field label."""
    root_getter = root_getter or _get_visible_new_enterprise_dialog
    root = root_getter(driver)
    if root is None:
        return None

    try:
        return driver.execute_script(
            """
            const root = arguments[0];
            const rawLabels = arguments[1] || [];
            const preferTextarea = Boolean(arguments[2]);
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const cleanLabel = (value) => normalize(value)
              .replace(/[\\*＊:：]/g, '')
              .trim();
            const expected = rawLabels.map(cleanLabel).filter(Boolean);
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
            const textMatches = (text) => {
              const cleaned = cleanLabel(text);
              if (!cleaned || cleaned.length > 80) return false;
              return expected.some((item) => cleaned === item || (cleaned.includes(item) && cleaned.length <= item.length + 8));
            };
            const controlSelector = 'input, textarea, [contenteditable="true"]';
            const usableControl = (el) => {
              if (!el || !visible(el) || el.disabled || el.readOnly) return false;
              const tag = String(el.tagName || '').toLowerCase();
              const type = String(el.getAttribute('type') || '').toLowerCase();
              if (tag === 'input' && ['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) return false;
              const cls = String(el.className || '').toLowerCase();
              if (/checkbox|radio|switch/.test(cls)) return false;
              return tag === 'input' || tag === 'textarea' || el.isContentEditable;
            };
            const addControlsFrom = (container, score, labelRect, out) => {
              if (!container || !visible(container)) return;
              for (const control of Array.from(container.querySelectorAll(controlSelector))) {
                if (!usableControl(control)) continue;
                const rect = control.getBoundingClientRect();
                if (labelRect && rect.left + rect.width < labelRect.left) continue;
                let finalScore = score;
                if (preferTextarea && String(control.tagName || '').toLowerCase() === 'textarea') finalScore += 25;
                if (!preferTextarea && String(control.tagName || '').toLowerCase() === 'input') finalScore += 10;
                out.push({ element: control, score: finalScore, top: rect.top, left: rect.left });
              }
            };
            const candidates = [];
            const labelSelector = [
              '.ui-formItem-label-span',
              '.ui-formItem-label',
              '.ui-form-item-label',
              '.ant-form-item-label',
              '.el-form-item__label',
              'label',
              'span',
              'div',
              'td',
              'th',
              '[title]',
              '[aria-label]'
            ].join(',');
            for (const label of Array.from(root.querySelectorAll(labelSelector))) {
              if (!visible(label)) continue;
              const rawText = normalize(
                label.innerText
                || label.textContent
                || label.getAttribute('title')
                || label.getAttribute('aria-label')
              );
              if (!textMatches(rawText)) continue;
              const labelRect = label.getBoundingClientRect();
              let node = label;
              for (let depth = 0; node && node !== root && depth < 7; depth += 1, node = node.parentElement) {
                const cls = String(node.className || '');
                const tag = String(node.tagName || '').toLowerCase();
                const looksLikeItem = /form|Form|field|Field|row|Row|col|Col|item|Item/.test(cls) || ['tr', 'td', 'li'].includes(tag);
                if (!looksLikeItem && depth < 2) continue;
                addControlsFrom(node, 100 - depth * 3, labelRect, candidates);
                if (candidates.length) break;
              }
            }

            // Geometry fallback: choose a control whose nearest same-row label on the left matches.
            for (const control of Array.from(root.querySelectorAll(controlSelector))) {
              if (!usableControl(control)) continue;
              const controlRect = control.getBoundingClientRect();
              for (const label of Array.from(root.querySelectorAll(labelSelector))) {
                if (!visible(label) || label === control || label.contains(control) || control.contains(label)) continue;
                const rawText = normalize(label.innerText || label.textContent || label.getAttribute('title') || label.getAttribute('aria-label'));
                if (!textMatches(rawText)) continue;
                const labelRect = label.getBoundingClientRect();
                const labelCenterY = labelRect.top + labelRect.height / 2;
                const controlCenterY = controlRect.top + controlRect.height / 2;
                const verticalDistance = Math.abs(labelCenterY - controlCenterY);
                const horizontallyBefore = labelRect.right <= controlRect.left + 32 && labelRect.right >= controlRect.left - 320;
                if (!horizontallyBefore || verticalDistance > Math.max(28, controlRect.height)) continue;
                let score = 75 - verticalDistance + Math.max(0, labelRect.right - controlRect.left) / 20;
                if (preferTextarea && String(control.tagName || '').toLowerCase() === 'textarea') score += 25;
                candidates.push({ element: control, score, top: controlRect.top, left: controlRect.left });
              }
            }

            // Attribute fallback for conventional placeholders/names.
            for (const control of Array.from(root.querySelectorAll(controlSelector))) {
              if (!usableControl(control)) continue;
              const haystack = [
                control.getAttribute('placeholder'),
                control.getAttribute('name'),
                control.getAttribute('title'),
                control.getAttribute('aria-label')
              ].map(normalize).join(' ');
              if (!textMatches(haystack)) continue;
              const rect = control.getBoundingClientRect();
              let score = 55;
              if (preferTextarea && String(control.tagName || '').toLowerCase() === 'textarea') score += 25;
              candidates.push({ element: control, score, top: rect.top, left: rect.left });
            }

            candidates.sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            root,
            labels,
            prefer_textarea,
        )
    except Exception:
        return None


def _set_text_control_value(driver: webdriver.Chrome, element: Any, value: str) -> bool:
    """Fill a text-like control and verify its value."""
    try:
        _safe_click(driver, element)
        try:
            element.clear()
        except Exception:
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(Keys.BACKSPACE)
        element.send_keys(value)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const tag = String(input.tagName || '').toLowerCase();
            if (tag === 'textarea') {
              const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
              if (setter) setter.call(input, value); else input.value = value;
            } else if (tag === 'input') {
              const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
              if (setter) setter.call(input, value); else input.value = value;
            } else if (input.isContentEditable) {
              input.textContent = value;
            } else {
              input.value = value;
            }
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            try {
              input.dispatchEvent(new InputEvent('input', {
                bubbles: true,
                cancelable: true,
                inputType: 'insertText',
                data: value
              }));
            } catch (e) {}
            input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter' }));
            input.dispatchEvent(new Event('blur', { bubbles: true }));
            """,
            element,
            value,
        )
        current_value = str(
            driver.execute_script(
                "return arguments[0].value || arguments[0].textContent || '';",
                element,
            )
            or ""
        ).strip()
        return current_value == value
    except Exception:
        return False


def _fill_enterprise_text_field(
    driver: webdriver.Chrome,
    display_name: str,
    labels: list[str],
    value: str,
    *,
    prefer_textarea: bool = False,
    root_getter=None,
) -> str:
    """Fill one labeled text field inside the new-enterprise dialog."""
    try:
        control = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_text_control_by_label(
                drv,
                labels,
                prefer_textarea=prefer_textarea,
                root_getter=root_getter,
            )
        )
    except TimeoutException:
        return f"未找到弹窗/页面中的「{display_name}」输入框。"

    if _set_text_control_value(driver, control, value):
        return f"{display_name}={value}"
    return f"填写「{display_name}」失败。"


def _fill_enterprise_form(driver: webdriver.Chrome, data: EnterpriseData) -> str:
    """Step 5: fill 企业名称、企业全称、备注信息."""
    fields = [
        (
            "企业名称",
            ["企业名称", "Enterprise Name", "Company Name"],
            data.name,
            False,
        ),
        (
            "企业全称",
            ["企业全称", "Enterprise Full Name", "Full Enterprise Name", "Company Full Name"],
            data.full_name,
            False,
        ),
        (
            "备注信息",
            ["备注信息", "备注", "Remark", "Remarks", "Notes"],
            data.remark,
            True,
        ),
    ]

    filled: list[str] = []
    failed: list[str] = []
    for display_name, labels, value, prefer_textarea in fields:
        result = _fill_enterprise_text_field(
            driver,
            display_name,
            labels,
            value,
            prefer_textarea=prefer_textarea,
        )
        if result.startswith(f"{display_name}="):
            filled.append(result)
        else:
            failed.append(result)

    if failed:
        return "填写新建企业表单失败：" + "；".join(failed)
    return "已填写新建企业表单：" + "；".join(filled) + "。"


def _click_save_new_enterprise(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new-enterprise dialog/drawer."""
    try:
        dialog = _get_visible_new_enterprise_dialog(driver)
        if dialog is None:
            return ""
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
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """,
                dialog,
            )
        ).strip()
    except Exception:
        return ""


def _enterprise_name_visible_in_body(driver: webdriver.Chrome, enterprise_name: str) -> bool:
    return enterprise_name in _body_text(driver)


def _save_new_enterprise_and_wait(driver: webdriver.Chrome, enterprise_name: str) -> str:
    """Save the enterprise and wait for success, duplicate, or validation feedback."""
    clicked = _click_save_new_enterprise(driver)
    if not clicked:
        return "未找到新建企业弹窗中的保存按钮。"

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
                return f"企业 {enterprise_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                return f"已保存企业 {enterprise_name}，页面提示：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                return f"已点击 {clicked}，但页面提示校验/保存失败：{joined_messages}。"

        body_text = _body_text(driver)
        if any(marker in body_text for marker in duplicate_markers) and enterprise_name in body_text:
            return f"企业 {enterprise_name} 已存在，视为通过。"
        if _enterprise_name_visible_in_body(driver, enterprise_name) and not _new_enterprise_dialog_is_open(driver):
            return f"已保存企业 {enterprise_name}，页面列表中已可见。"
        if not _new_enterprise_dialog_is_open(driver):
            # The drawer closed without a toast we can read; continue with list/search verification.
            break
        time.sleep(0.4)

    if _enterprise_name_visible_in_body(driver, enterprise_name):
        return f"已保存企业 {enterprise_name}，页面已显示该企业。"

    # Some runs close the drawer and refresh the left tree/list without a toast
    # that Selenium can read. In that case, actively search the external-org
    # tree by the new enterprise name before treating the save as failed.
    try:
        WebDriverWait(driver, 10, poll_frequency=0.5).until(
            lambda drv: not _new_enterprise_dialog_is_open(drv)
        )
        if _search_enterprise_by_name(driver, enterprise_name):
            return f"已保存企业 {enterprise_name}，保存后已通过搜索验证可见。"
    except TimeoutException:
        pass

    return (
        f"已点击 {clicked}，但未在页面上确认企业 {enterprise_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def _find_enterprise_search_input(driver: webdriver.Chrome):
    """Return an enterprise list search input outside any visible dialog/drawer."""
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
            const insideVisibleDialog = (el) => {
              const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
              return dialog && visible(dialog);
            };
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const candidates = [];
            for (const input of Array.from(document.querySelectorAll('input'))) {
              if (!visible(input) || insideVisibleDialog(input) || input.disabled || input.readOnly) continue;
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (!['', 'text', 'search'].includes(type)) continue;
              const placeholder = normalize(input.getAttribute('placeholder') || input.getAttribute('aria-label') || input.getAttribute('title'));
              if (!/企业|Enterprise|Company/i.test(placeholder)) continue;
              const rect = input.getBoundingClientRect();
              candidates.push({ element: input, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) => a.top - b.top || b.left - a.left);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _search_enterprise_by_name(driver: webdriver.Chrome, enterprise_name: str) -> bool:
    """Search the external enterprise list by enterprise name when a search box is available."""
    if _enterprise_name_visible_in_body(driver, enterprise_name):
        return True

    search_input = _find_enterprise_search_input(driver)
    if search_input is None:
        return _enterprise_name_visible_in_body(driver, enterprise_name)

    try:
        _safe_click(driver, search_input)
        try:
            search_input.clear()
        except Exception:
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.BACKSPACE)
        search_input.send_keys(enterprise_name)
        search_input.send_keys(Keys.ENTER)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
            if (setter) setter.call(input, value); else input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.dispatchEvent(new KeyboardEvent('keydown', {
              bubbles: true,
              cancelable: true,
              key: 'Enter',
              code: 'Enter'
            }));
            input.dispatchEvent(new KeyboardEvent('keyup', {
              bubbles: true,
              cancelable: true,
              key: 'Enter',
              code: 'Enter'
            }));
            """,
            search_input,
            enterprise_name,
        )
        try:
            WebDriverWait(driver, 6, poll_frequency=0.4).until(
                lambda drv: _enterprise_name_visible_in_body(drv, enterprise_name)
            )
            return True
        except TimeoutException:
            return _enterprise_name_visible_in_body(driver, enterprise_name)
    except Exception:
        return _enterprise_name_visible_in_body(driver, enterprise_name)


def _click_enterprise_tree_node(driver: webdriver.Chrome, enterprise_name: str) -> str:
    """Click a specific enterprise node in the external organization tree."""
    try:
        return str(
            driver.execute_script(
                """
                const enterpriseName = arguments[0];
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
                for (const raw of Array.from(document.body.querySelectorAll('.ui-tree-node-content, .ui-tree-node, .ui-tree-bar, li, span, div'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title'));
                  if (text !== enterpriseName) continue;
                  const clickable = raw.closest('.ui-tree-node-content, .ui-tree-node, .ui-tree-bar, li') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  // External-org tree is the middle-left column, not the header/content area.
                  if (rect.left < 200 || rect.left > 520 || rect.top < 90) continue;
                  candidates.push({
                    element: clickable,
                    text,
                    left: rect.left,
                    top: rect.top,
                    className: String(clickable.className || '')
                  });
                }
                candidates.sort((a, b) =>
                  (b.className.includes('ui-tree-node-content') ? 1 : 0)
                    - (a.className.includes('ui-tree-node-content') ? 1 : 0)
                  || a.top - b.top
                  || b.left - a.left
                );
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent) || enterpriseName;
                fireClick(target);
                return clickedText;
                """,
                enterprise_name,
            )
        ).strip()
    except Exception:
        return ""


def _enterprise_page_is_open(driver: webdriver.Chrome, enterprise_name: str) -> bool:
    """Return True when the newly-created enterprise page is selected."""
    try:
        body_text = _body_text(driver)
        return (
            f"企业：{enterprise_name}" in body_text
            and "新建部门" in body_text
        )
    except Exception:
        return False


def _open_enterprise_page(driver: webdriver.Chrome, enterprise_name: str) -> str:
    """
    Select the newly-created enterprise node.

    The page already shows a top-right 新建部门 button after the enterprise is
    selected, so this helper deliberately does not switch to the 下级部门 tab.
    """
    if _enterprise_page_is_open(driver, enterprise_name):
        return f"已在企业 {enterprise_name} 页面，并已看到右上角「新建部门」按钮。"

    if not _search_enterprise_by_name(driver, enterprise_name):
        return f"未在外部组织维护树/列表中找到企业 {enterprise_name}。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_enterprise_tree_node(driver, enterprise_name)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                lambda drv: _enterprise_page_is_open(drv, enterprise_name)
            )
            return f"已点击企业节点 {last_clicked}，并已进入该企业页面。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        f"未能进入企业 {enterprise_name} 页面："
        f"{'已点击 ' + last_clicked + ' 但未看到企业页面右上角「新建部门」。' if last_clicked else '未找到可点击的企业节点。'}"
    )


def _new_external_department_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the 新建部门 dialog/drawer is visible in external org page."""
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
                  if (text.includes('新建部门') && (text.includes('部门名称') || text.includes('名称'))) return true;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _get_visible_new_external_department_dialog(driver: webdriver.Chrome):
    """Return the visible 新建部门 dialog/drawer WebElement, if present."""
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
                return text.includes('新建部门') && (text.includes('部门名称') || text.includes('名称'));
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


def _click_new_external_department_button(driver: webdriver.Chrome) -> str:
    """Click the top-right 新建部门 button on the selected enterprise page."""
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
                const insideVisibleDialog = (el) => {
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (text !== '新建部门' && text !== '新增部门' && text !== 'New Department' && text !== 'Add Department') continue;
                  const rect = el.getBoundingClientRect();
                  // The user clarified this is the enterprise page top-right button.
                  if (rect.left < 500 || rect.top < 40 || rect.top > 140) continue;
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.x - a.x || a.y - b.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_new_external_department_dialog(
    driver: webdriver.Chrome,
    enterprise_name: str,
) -> str:
    """Click the enterprise page 新建部门 button and verify the dialog opens."""
    if not _enterprise_page_is_open(driver, enterprise_name):
        return f"未在企业 {enterprise_name} 页面，无法点击「新建部门」。"

    if _new_external_department_dialog_is_open(driver):
        return "新建部门弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_external_department_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_external_department_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建部门弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建部门弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到企业页面右上角可点击的「新建部门」按钮。'}"
    )


def _fill_optional_external_department_text_field(
    driver: webdriver.Chrome,
    display_name: str,
    labels: list[str],
    value: str,
    *,
    prefer_textarea: bool = False,
) -> str:
    """Fill an optional text field if it exists in the external department form."""
    control = _find_text_control_by_label(
        driver,
        labels,
        prefer_textarea=prefer_textarea,
        root_getter=_get_visible_new_external_department_dialog,
    )
    if control is None:
        return f"{display_name}字段未显示，已跳过"
    if _set_text_control_value(driver, control, value):
        return f"{display_name}={value}"
    return f"填写「{display_name}」失败"


def _fill_external_department_form(
    driver: webdriver.Chrome,
    data: ExternalDepartmentData,
) -> str:
    """Fill required and common fields in the external new-department form."""
    name_result = _fill_enterprise_text_field(
        driver,
        "部门名称",
        ["部门名称", "名称", "Department Name", "Name"],
        data.name,
        root_getter=_get_visible_new_external_department_dialog,
    )
    if not name_result.startswith("部门名称="):
        return "填写新建部门表单失败：" + name_result

    filled = [name_result]
    for result in [
        _fill_optional_external_department_text_field(
            driver,
            "部门全称",
            ["部门全称", "全称", "Department Full Name", "Full Name"],
            data.full_name,
        ),
        _fill_optional_external_department_text_field(
            driver,
            "备注信息",
            ["备注信息", "备注", "Remark", "Remarks", "Notes"],
            data.remark,
            prefer_textarea=True,
        ),
    ]:
        if result.endswith("失败"):
            return "填写新建部门表单失败：" + result
        filled.append(result)

    return "已填写新建部门表单：" + "；".join(filled) + "。"


def _click_save_new_external_department(driver: webdriver.Chrome) -> str:
    """Click 保存 in the external new-department dialog/drawer."""
    try:
        dialog = _get_visible_new_external_department_dialog(driver)
        if dialog is None:
            return ""
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
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """
                ,
                dialog,
            )
        ).strip()
    except Exception:
        return ""


def _save_new_external_department_and_wait(
    driver: webdriver.Chrome,
    department_name: str,
) -> str:
    """Save the external department and wait for success/validation feedback."""
    clicked = _click_save_new_external_department(driver)
    if not clicked:
        return "未找到新建部门弹窗中的保存按钮。"

    duplicate_markers = ("已存在", "重复", "duplicate", "already exists", "Already exists")
    success_markers = ("保存成功", "新增成功", "创建成功", "操作成功", "success", "Success")
    validation_markers = ("必填", "不能为空", "请输入", "校验", "失败", "错误", "error", "Error")
    deadline = time.monotonic() + 15
    last_messages: list[str] = []
    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"部门 {department_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                return f"已保存部门 {department_name}，页面提示：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                return f"已点击 {clicked}，但页面提示校验/保存失败：{joined_messages}。"

        body_text = _body_text(driver)
        if any(marker in body_text for marker in duplicate_markers) and department_name in body_text:
            return f"部门 {department_name} 已存在，视为通过。"
        if department_name in body_text and not _new_external_department_dialog_is_open(driver):
            return f"已保存部门 {department_name}，页面列表/树中已可见。"
        if not _new_external_department_dialog_is_open(driver):
            break
        time.sleep(0.4)

    if department_name in _body_text(driver):
        return f"已保存部门 {department_name}，页面已显示该部门。"

    return (
        f"已点击 {clicked}，但未在页面上确认部门 {department_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def _external_department_page_is_open(
    driver: webdriver.Chrome,
    department_name: str,
) -> bool:
    """Return True when the external department page is selected."""
    try:
        body_text = _body_text(driver)
        return (
            f"部门：{department_name}" in body_text
            and "联系人" in body_text
            and "新建" in body_text
        )
    except Exception:
        return False


def _open_external_department_page(
    driver: webdriver.Chrome,
    department_name: str,
    enterprise_name: str = "",
) -> str:
    """Search/select the new external department in the left tree."""
    if _external_department_page_is_open(driver, department_name):
        return f"已在部门 {department_name} 页面。"

    last_clicked = ""
    last_search_ok = False
    # Newly-created department nodes can take a few seconds to appear in the
    # external-org tree/search index after the save toast. Retry the tree search
    # before treating the node as missing.
    for attempt in range(8):
        if enterprise_name and attempt:
            _open_enterprise_page(driver, enterprise_name)
            time.sleep(0.8)

        # Reuse the external-org tree search input. Despite the function name,
        # it searches the same tree by enterprise or department name.
        last_search_ok = _search_enterprise_by_name(driver, department_name)
        if not last_search_ok:
            time.sleep(1.5)
            continue

        for _ in range(3):
            last_clicked = _click_enterprise_tree_node(driver, department_name)
            if not last_clicked:
                time.sleep(0.5)
                continue
            try:
                WebDriverWait(driver, 10, poll_frequency=0.4).until(
                    lambda drv: _external_department_page_is_open(drv, department_name)
                )
                return f"已点击部门节点 {last_clicked}，并已进入该部门页面。"
            except TimeoutException:
                time.sleep(0.5)

    return (
        f"未能进入部门 {department_name} 页面："
        f"{'已点击 ' + last_clicked + ' 但未看到部门页面。' if last_clicked else ('搜索到部门但未找到可点击节点。' if last_search_ok else '多次搜索仍未找到部门节点。')}"
    )


def _external_contact_tab_is_active(driver: webdriver.Chrome) -> bool:
    """Return True when the right-side 联系人 tab is active."""
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
                const isActive = (el) => {
                  const attrs = [
                    el.className,
                    el.getAttribute('aria-selected'),
                    el.getAttribute('aria-current'),
                    el.getAttribute('data-active')
                  ].map(normalize).join(' ').toLowerCase();
                  return /active|selected|current|checked|true/.test(attrs);
                };
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (text !== '联系人') continue;
                  const tab = raw.closest('.ui-menu-list-item, [role="tab"], li, button, a') || raw;
                  if (!visible(tab)) continue;
                  const rect = tab.getBoundingClientRect();
                  if (rect.left < 500 || rect.left > 900 || rect.top < 50 || rect.top > 130) continue;
                  if (isActive(tab)) return true;
                }
                const bodyText = normalize(document.body.innerText || document.body.textContent);
                return bodyText.includes('姓名 联系方式 所属企业 部门 岗位')
                  && bodyText.includes('批量操作');
                """
            )
        )
    except Exception:
        return False


def _click_external_contact_tab(driver: webdriver.Chrome) -> str:
    """Click the department page right-side 联系人 tab."""
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
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('.ui-menu-list-item, .ui-menu-parent-icon, [role="tab"], li, button, a, span, div'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (text !== '联系人') continue;
                  const clickable = raw.closest('.ui-menu-list-item, .ui-menu-parent-icon, [role="tab"], li, button, a') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  if (rect.left < 500 || rect.left > 900 || rect.top < 50 || rect.top > 130) continue;
                  candidates.push({ element: clickable, text, left: rect.left, top: rect.top });
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


def _open_external_contact_tab(
    driver: webdriver.Chrome,
    department_name: str,
) -> str:
    """Switch the selected department page to the 联系人 tab."""
    if not _external_department_page_is_open(driver, department_name):
        return f"未在部门 {department_name} 页面，无法切换「联系人」tab。"

    if _external_contact_tab_is_active(driver):
        return "已在「联系人」tab。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_external_contact_tab(driver)
        if not last_clicked:
            time.sleep(0.4)
            continue
        try:
            WebDriverWait(driver, 8, poll_frequency=0.3).until(
                _external_contact_tab_is_active
            )
            return f"已点击 {last_clicked} tab，并已进入联系人列表。"
        except TimeoutException:
            time.sleep(0.4)

    return (
        "未能进入「联系人」tab："
        f"{'已点击 ' + last_clicked + ' 但未验证成功。' if last_clicked else '未找到可点击的联系人tab。'}"
    )


def _new_external_contact_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the 新建外部联系人 drawer is visible."""
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
                  if (text.includes('新建外部联系人') && text.includes('姓名')) return true;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _get_visible_new_external_contact_dialog(driver: webdriver.Chrome):
    """Return the visible 新建外部联系人 drawer WebElement, if present."""
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
                return text.includes('新建外部联系人') && text.includes('姓名');
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


def _click_new_external_contact_button(driver: webdriver.Chrome) -> str:
    """Click the 联系人 tab toolbar 新建 button."""
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
                const insideVisibleDialog = (el) => {
                  const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap, .ui-modal, .modal, .ant-modal, .el-dialog, [role="dialog"]');
                  return dialog && visible(dialog);
                };
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button.ui-btn, button, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (text !== '新建' && text !== '新增' && text !== 'New' && text !== 'Add') continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 500 || rect.top < 40 || rect.top > 160) continue;
                  const tag = String(el.tagName || '').toUpperCase();
                  const className = String(el.className || '');
                  const score = (tag === 'BUTTON' ? 100 : 0) + rect.left / 100;
                  candidates.push({ element: el, text, x: rect.left, y: rect.top, score, className });
                }
                candidates.sort((a, b) => b.score - a.score || b.x - a.x || a.y - b.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_new_external_contact_dialog(driver: webdriver.Chrome) -> str:
    """Click 新建 in 联系人 tab and verify the new-contact drawer opens."""
    if not _external_contact_tab_is_active(driver):
        return "未在「联系人」tab，无法点击「新建」。"

    if _new_external_contact_dialog_is_open(driver):
        return "新建外部联系人弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_external_contact_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_external_contact_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建外部联系人弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建外部联系人弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到联系人tab中的「新建」按钮。'}"
    )


def _fill_optional_external_contact_text_field(
    driver: webdriver.Chrome,
    display_name: str,
    labels: list[str],
    value: str,
    *,
    prefer_textarea: bool = False,
) -> str:
    """Fill an optional text field if it exists in the external contact form."""
    control = _find_text_control_by_label(
        driver,
        labels,
        prefer_textarea=prefer_textarea,
        root_getter=_get_visible_new_external_contact_dialog,
    )
    if control is None:
        return f"{display_name}字段未显示，已跳过"
    if _set_text_control_value(driver, control, value):
        return f"{display_name}={value}"
    return f"填写「{display_name}」失败"


def _find_external_contact_name_input(driver: webdriver.Chrome):
    """Return the visible input immediately to the right of 新建外部联系人「姓名」."""
    dialog = _get_visible_new_external_contact_dialog(driver)
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
            const candidates = [];
            const labels = Array.from(dialog.querySelectorAll(
              '.ui-formItem-label-span, .ui-formItem-label, label, span, div'
            ));
            for (const label of labels) {
              if (!visible(label)) continue;
              if (cleanLabel(label.innerText || label.textContent || label.getAttribute('title')) !== '姓名') continue;
              const labelRect = label.getBoundingClientRect();
              // The name field is in 基本资料 near the top of the drawer.
              if (labelRect.top < 120 || labelRect.top > 230 || labelRect.left < 640 || labelRect.left > 820) continue;

              let container = label.closest('.ui-formItem')
                || label.closest('.ui-form-col')
                || label.closest('.ui-layout-col')
                || label.parentElement;
              for (let depth = 0; container && container !== dialog && depth < 5; depth += 1, container = container.parentElement) {
                for (const input of Array.from(container.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
                  if (!visible(input) || input.disabled || input.readOnly) continue;
                  const tag = String(input.tagName || '').toLowerCase();
                  const type = String(input.getAttribute('type') || '').toLowerCase();
                  if (tag === 'input' && ['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) continue;
                  const rect = input.getBoundingClientRect();
                  if (rect.top < labelRect.top - 10 || rect.top > labelRect.top + 25) continue;
                  if (rect.left <= labelRect.right || rect.left > 1050) continue;
                  candidates.push({
                    element: input,
                    score: 100 - Math.abs(rect.top - labelRect.top) - Math.abs(rect.left - 780) / 20,
                    top: rect.top,
                    left: rect.left
                  });
                }
                if (candidates.length) break;
              }
            }
            candidates.sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """,
            dialog,
        )
    except Exception:
        return None


def _fill_external_contact_name_field(
    driver: webdriver.Chrome,
    contact_name: str,
) -> str:
    """Fill and strictly verify the required 新建外部联系人「姓名」field."""
    try:
        name_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_external_contact_name_input(drv)
        )
    except TimeoutException:
        return "未找到新建外部联系人弹窗中的「姓名」输入框。"

    if not _set_text_control_value(driver, name_input, contact_name):
        return "填写「姓名」失败：输入框未保持目标值。"

    actual = str(name_input.get_attribute("value") or "").strip()
    if actual != contact_name:
        return f"填写「姓名」失败：当前值为 {actual or '空'}。"
    return f"姓名={contact_name}"


def _fill_external_contact_form(
    driver: webdriver.Chrome,
    data: ExternalContactData,
) -> str:
    """Fill required contact fields in the external contact form."""
    name_result = _fill_external_contact_name_field(driver, data.name)
    if not name_result.startswith("姓名="):
        return "填写新建外部联系人表单失败：" + name_result

    return "已填写新建外部联系人必填信息：" + name_result + "。"


def _click_save_new_external_contact(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new external contact drawer."""
    try:
        dialog = _get_visible_new_external_contact_dialog(driver)
        if dialog is None:
            return ""
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
                  if (text !== '保存' && text !== 'Save') continue;
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
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """,
                dialog,
            )
        ).strip()
    except Exception:
        return ""


def _find_external_contact_search_input(driver: webdriver.Chrome):
    """Return the right-top search input in the external contact list."""
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
            const contactPageText = normalize(document.body.innerText || document.body.textContent);
            if (!contactPageText.includes('联系人')) return null;

            const newContactButtons = [];
            for (const raw of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
              if (!visible(raw) || insideVisibleDialog(raw)) continue;
              const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title') || raw.getAttribute('aria-label'));
              if (!['新建', '新增', 'New', 'Add'].includes(text)) continue;
              const rect = raw.getBoundingClientRect();
              if (rect.left < 500 || rect.top < 40 || rect.top > 140) continue;
              newContactButtons.push({ rect });
            }
            newContactButtons.sort((a, b) => b.rect.left - a.rect.left || a.rect.top - b.rect.top);
            const buttonRect = newContactButtons[0]?.rect || null;

            const candidates = [];
            for (const input of Array.from(document.querySelectorAll('input'))) {
              if (!visible(input) || insideVisibleDialog(input) || input.disabled || input.readOnly) continue;
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (!['', 'text', 'search'].includes(type)) continue;
              const rect = input.getBoundingClientRect();
              // Exclude the external-org tree search on the left; the contact
              // search box is in the right-side toolbar above the contact list.
              if (rect.left < 520 || rect.top < 40 || rect.top > 180) continue;
              const placeholder = normalize(
                input.getAttribute('placeholder')
                || input.getAttribute('aria-label')
                || input.getAttribute('title')
              );
              const wrapper = input.closest('.ui-input, .ui-search, .ui-input-wrap, .ui-input-wrapper, .ui-formItem, div')
                || input.parentElement;
              const wrapperText = normalize(wrapper?.innerText || wrapper?.textContent || '');
              const attrs = [
                input.className,
                input.id,
                input.name,
                placeholder,
                wrapper?.className,
                wrapperText
              ].map(normalize).join(' ');

              let score = rect.left / 10 + Math.max(0, 260 - rect.top) / 4;
              if (/请输入\\s*姓名/.test(placeholder)) score += 1000;
              else if (/姓名/.test(placeholder)) score += 500;
              if (buttonRect) {
                const inputCenterY = rect.top + rect.height / 2;
                const buttonCenterY = buttonRect.top + buttonRect.height / 2;
                const sameRow = Math.abs(inputCenterY - buttonCenterY) <= 80;
                if (sameRow) score += 250;
                if (sameRow && rect.left >= buttonRect.right - 20) score += 200;
                score -= Math.abs(rect.left - buttonRect.right) / 5;
              }
              if (/联系人|外部联系人|contact/i.test(attrs)) score += 120;
              if (/搜索|查询|search/i.test(attrs)) score += 90;
              if (/请输入/.test(placeholder)) score += 40;
              if (rect.left > window.innerWidth * 0.62) score += 60;
              if (rect.width >= 120) score += 20;

              candidates.push({ element: input, score, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) => b.score - a.score || b.left - a.left || a.top - b.top);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _external_contact_name_visible_in_list(
    driver: webdriver.Chrome,
    contact_name: str,
) -> bool:
    """Return True when the contact name is visible in the right-side contact list."""
    try:
        return bool(
            driver.execute_script(
                """
                const contactName = arguments[0];
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
                const ignoredTags = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'SVG', 'PATH', 'INPUT', 'TEXTAREA']);
                const selectors = [
                  'td',
                  '.ui-table-cell',
                  '.ui-table-row',
                  '.ui-list-item',
                  '[role="row"]',
                  '[role="cell"]',
                  'a',
                  'span',
                  'div'
                ].join(',');
                for (const el of Array.from(document.body.querySelectorAll(selectors))) {
                  if (ignoredTags.has(el.tagName) || !visible(el) || insideVisibleDialog(el)) continue;
                  const rect = el.getBoundingClientRect();
                  // Right-side list only. Avoid matching the left organization
                  // tree or broad page containers.
                  if (rect.left < 500 || rect.top < 110) continue;
                  if (rect.width > window.innerWidth * 0.9 && rect.height > window.innerHeight * 0.65) continue;
                  const text = normalize(
                    el.innerText
                    || el.textContent
                    || el.getAttribute('title')
                    || el.getAttribute('aria-label')
                  );
                  if (text && text.includes(contactName)) return true;
                }
                return false;
                """,
                contact_name,
            )
        )
    except Exception:
        return False


def _click_external_contact_search_icon(driver: webdriver.Chrome, search_input: Any) -> str:
    """Click the search icon/button near the contact-list search input, if present."""
    try:
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

                const inputRect = input.getBoundingClientRect();
                const roots = [];
                let root = input.parentElement;
                for (let depth = 0; root && depth < 6; depth += 1, root = root.parentElement) {
                  roots.push(root);
                }

                const candidates = [];
                for (const rootEl of roots) {
                  for (const el of Array.from(rootEl.querySelectorAll('button, a, [role="button"], i, span, svg, .ui-icon, .iconfont'))) {
                    if (el === input || !visible(el)) continue;
                    const rect = el.getBoundingClientRect();
                    if (rect.left < inputRect.left || rect.left > inputRect.right + 80) continue;
                    if (Math.abs((rect.top + rect.height / 2) - (inputRect.top + inputRect.height / 2)) > 28) continue;
                    const attrs = [
                      el.innerText,
                      el.textContent,
                      el.className,
                      el.getAttribute('title'),
                      el.getAttribute('aria-label'),
                      el.getAttribute('name')
                    ].map(normalize).join(' ');
                    let score = 10;
                    if (/搜索|查询|search|magnif|sousuo/i.test(attrs)) score += 120;
                    if (rect.left >= inputRect.right - 45) score += 60;
                    candidates.push({ element: el, score, text: normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label')) });
                  }
                  if (candidates.length) break;
                }

                candidates.sort((a, b) => b.score - a.score);
                const target = candidates[0]?.element;
                if (!target) return '';
                fireClick(target);
                return candidates[0].text || '搜索图标';
                """,
                search_input,
            )
        ).strip()
    except Exception:
        return ""


def _search_external_contact_by_name(
    driver: webdriver.Chrome,
    contact_name: str,
) -> str:
    """Search the contact list by contact name and verify the result is visible."""
    try:
        WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: not _new_external_contact_dialog_is_open(drv)
        )
    except TimeoutException:
        return "新建外部联系人弹窗仍未关闭，无法在联系人列表右上角搜索验证。"

    try:
        search_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_external_contact_search_input(drv)
        )
    except TimeoutException:
        return f"未找到联系人列表右上角搜索输入框，无法搜索验证外部联系人 {contact_name}。"

    if not _set_text_control_value(driver, search_input, contact_name):
        return f"填写联系人搜索框失败，无法搜索验证外部联系人 {contact_name}。"

    try:
        input_label = str(
            driver.execute_script(
                """
                const el = arguments[0];
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                return normalize(
                  el?.getAttribute('placeholder')
                  || el?.getAttribute('aria-label')
                  || el?.getAttribute('title')
                  || ''
                );
                """,
                search_input,
            )
            or ""
        ).strip()
        input_value = str(
            driver.execute_script(
                "return arguments[0]?.value || arguments[0]?.textContent || '';",
                search_input,
            )
            or ""
        ).strip()
    except Exception:
        input_label = ""
        input_value = ""

    if contact_name not in input_value:
        return (
            "未能确认已在联系人右上角搜索框中输入外部联系人姓名"
            f"（当前搜索框值：{input_value or '空'}）。"
        )

    try:
        _safe_click(driver, search_input)
        search_input.send_keys(Keys.ENTER)
    except Exception:
        pass

    try:
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
            if (setter) setter.call(input, value); else input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            for (const type of ['keydown', 'keypress', 'keyup']) {
              input.dispatchEvent(new KeyboardEvent(type, {
                bubbles: true,
                cancelable: true,
                key: 'Enter',
                code: 'Enter'
              }));
            }
            """,
            search_input,
            contact_name,
        )
    except Exception:
        pass

    clicked_icon = _click_external_contact_search_icon(driver, search_input)
    trigger_text = f"并点击 {clicked_icon}" if clicked_icon else "并按 Enter"
    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: _external_contact_name_visible_in_list(drv, contact_name)
        )
        return (
            f"已在联系人右上角搜索框{f'（{input_label}）' if input_label else ''}"
            f"输入 {contact_name}，{trigger_text}，"
            "搜索结果中已显示该外部联系人。"
        )
    except TimeoutException:
        messages = _collect_visible_messages(driver, fast=True)
        return (
            f"已在联系人右上角搜索框{f'（{input_label}）' if input_label else ''}"
            f"输入 {contact_name}，{trigger_text}，"
            "但搜索结果未显示该外部联系人"
            f"{'；页面提示：' + '；'.join(messages) if messages else ''}。"
        )


def _save_new_external_contact_and_wait(
    driver: webdriver.Chrome,
    contact_name: str,
) -> str:
    """Save the contact and wait for success/validation feedback."""
    clicked = _click_save_new_external_contact(driver)
    if not clicked:
        return "未找到新建外部联系人弹窗中的保存按钮。"

    duplicate_markers = ("已存在", "重复", "duplicate", "already exists", "Already exists")
    success_markers = ("保存成功", "新增成功", "创建成功", "操作成功", "success", "Success")
    validation_markers = ("必填", "不能为空", "请输入", "校验", "失败", "错误", "error", "Error")
    deadline = time.monotonic() + 15
    last_messages: list[str] = []
    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"外部联系人 {contact_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                return f"已保存外部联系人 {contact_name}，页面提示：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                return f"已点击 {clicked}，但页面提示校验/保存失败：{joined_messages}。"

        body_text = _body_text(driver)
        if any(marker in body_text for marker in duplicate_markers) and contact_name in body_text:
            return f"外部联系人 {contact_name} 已存在，视为通过。"
        if contact_name in body_text and not _new_external_contact_dialog_is_open(driver):
            return f"已保存外部联系人 {contact_name}，联系人列表中已可见。"
        if not _new_external_contact_dialog_is_open(driver):
            dialog_closed_without_error = True
            break
        time.sleep(0.4)

    if contact_name in _body_text(driver):
        return f"已保存外部联系人 {contact_name}，页面已显示该联系人。"

    if dialog_closed_without_error:
        return (
            f"已保存外部联系人 {contact_name}，新建外部联系人弹窗已关闭且未发现错误提示，"
            "将继续通过右上角搜索框验证。"
        )

    return (
        f"已点击 {clicked}，但未在页面上确认外部联系人 {contact_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def create_new_enterprise_with_department(
    driver: webdriver.Chrome,
    enterprise_name: str = "",
    enterprise_full_name: str = "",
    enterprise_remark: str = "",
    department_name: str = "",
    department_full_name: str = "",
    department_remark: str = "",
) -> str:
    """
    Create a new external enterprise, then create a department from that
    enterprise page's top-right 新建部门 button.
    """
    enterprise = _generate_enterprise_data(
        enterprise_name=enterprise_name,
        enterprise_full_name=enterprise_full_name,
        remark=enterprise_remark,
    )
    department = _generate_external_department_data(
        department_name=department_name,
        department_full_name=department_full_name,
        department_remark=department_remark,
    )

    enterprise_result = create_new_enterprise(
        driver,
        enterprise_name=enterprise.name,
        enterprise_full_name=enterprise.full_name,
        remark=enterprise.remark,
    )
    if not enterprise_result.startswith("已完成新建企业流程"):
        return (
            f"未能在新企业下创建部门 {department.name}："
            f"新建企业步骤未完成。{enterprise_result}"
        )

    step_results: list[str] = [f"步骤1-新建企业：{enterprise_result}"]

    open_enterprise_result = _open_enterprise_page(driver, enterprise.name)
    step_results.append(f"步骤2-进入新企业页面：{open_enterprise_result}")
    if not open_enterprise_result.startswith("已"):
        return f"未能在新企业 {enterprise.name} 下创建部门 {department.name}：" + " ".join(step_results)

    new_department_result = _open_new_external_department_dialog(driver, enterprise.name)
    step_results.append(f"步骤3-新建部门：{new_department_result}")
    if not new_department_result.startswith(("已", "新建部门弹窗")):
        return f"未能在新企业 {enterprise.name} 下创建部门 {department.name}：" + " ".join(step_results)

    fill_result = _fill_external_department_form(driver, department)
    step_results.append(f"步骤4-填写部门信息：{fill_result}")
    if not fill_result.startswith("已填写"):
        return f"未能在新企业 {enterprise.name} 下创建部门 {department.name}：" + " ".join(step_results)

    save_result = _save_new_external_department_and_wait(driver, department.name)
    step_results.append(f"步骤5-保存部门：{save_result}")
    if save_result.startswith(("已保存", "部门")):
        return (
            f"已完成新建企业并在该企业下新建部门流程："
            f"企业名称={enterprise.name}，部门名称={department.name}，"
            f"部门全称={department.full_name}，部门备注={department.remark}。"
            + " ".join(step_results)
        )

    return (
        f"未能确认在新企业 {enterprise.name} 下创建部门 {department.name} 成功："
        + " ".join(step_results)
    )


def create_new_enterprise_with_department_and_contact(
    driver: webdriver.Chrome,
    enterprise_name: str = "",
    enterprise_full_name: str = "",
    enterprise_remark: str = "",
    department_name: str = "",
    department_full_name: str = "",
    department_remark: str = "",
    contact_name: str = "",
    contact_mobile: str = "",
    contact_email: str = "",
    contact_phone: str = "",
    contact_remark: str = "",
) -> str:
    """
    Create a new external enterprise, create a department under it, then create
    an external contact from that department page's 联系人 tab.
    """
    enterprise = _generate_enterprise_data(
        enterprise_name=enterprise_name,
        enterprise_full_name=enterprise_full_name,
        remark=enterprise_remark,
    )
    department = _generate_external_department_data(
        department_name=department_name,
        department_full_name=department_full_name,
        department_remark=department_remark,
    )
    contact = _generate_external_contact_data(
        contact_name=contact_name,
        contact_mobile=contact_mobile,
        contact_email=contact_email,
        contact_phone=contact_phone,
        contact_remark=contact_remark,
    )

    department_flow_result = create_new_enterprise_with_department(
        driver,
        enterprise_name=enterprise.name,
        enterprise_full_name=enterprise.full_name,
        enterprise_remark=enterprise.remark,
        department_name=department.name,
        department_full_name=department.full_name,
        department_remark=department.remark,
    )
    if not department_flow_result.startswith("已完成新建企业并在该企业下新建部门流程"):
        return (
            f"未能在新部门下创建外部联系人 {contact.name}："
            f"新建企业/部门步骤未完成。{department_flow_result}"
        )

    step_results: list[str] = [f"步骤1-新建企业和部门：{department_flow_result}"]

    open_department_result = _open_external_department_page(
        driver,
        department.name,
        enterprise_name=enterprise.name,
    )
    step_results.append(f"步骤2-进入新部门页面：{open_department_result}")
    if not open_department_result.startswith("已"):
        return (
            f"未能在新部门 {department.name} 下创建外部联系人 {contact.name}："
            + " ".join(step_results)
        )

    contact_tab_result = _open_external_contact_tab(driver, department.name)
    step_results.append(f"步骤3-切换联系人tab：{contact_tab_result}")
    if not contact_tab_result.startswith("已"):
        return (
            f"未能在新部门 {department.name} 下创建外部联系人 {contact.name}："
            + " ".join(step_results)
        )

    new_contact_result = _open_new_external_contact_dialog(driver)
    step_results.append(f"步骤4-新建外部联系人：{new_contact_result}")
    if not new_contact_result.startswith(("已", "新建外部联系人弹窗")):
        return (
            f"未能在新部门 {department.name} 下创建外部联系人 {contact.name}："
            + " ".join(step_results)
        )

    fill_result = _fill_external_contact_form(driver, contact)
    step_results.append(f"步骤5-填写外部联系人必填信息：{fill_result}")
    if not fill_result.startswith("已填写"):
        return (
            f"未能在新部门 {department.name} 下创建外部联系人 {contact.name}："
            + " ".join(step_results)
        )

    save_result = _save_new_external_contact_and_wait(driver, contact.name)
    step_results.append(f"步骤6-保存外部联系人：{save_result}")

    if not save_result.startswith(("已保存", "外部联系人")):
        return (
            f"未能确认在新部门 {department.name} 下创建外部联系人 {contact.name} 成功："
            + " ".join(step_results)
        )

    search_result = _search_external_contact_by_name(driver, contact.name)
    step_results.append(f"步骤7-搜索验证外部联系人：{search_result}")
    if search_result.startswith("已"):
        return (
            f"已完成新建企业、新建部门并在该部门下新建外部联系人流程："
            f"企业名称={enterprise.name}，部门名称={department.name}，"
            f"外部联系人姓名={contact.name}。"
            + " ".join(step_results)
        )

    return (
        f"已保存但未能通过右上角搜索框确认外部联系人 {contact.name} 可检索："
        + " ".join(step_results)
    )


def create_new_enterprise(
    driver: webdriver.Chrome,
    enterprise_name: str = "",
    enterprise_full_name: str = "",
    remark: str = "",
) -> str:
    """
    Complete the Org. Structure -> 外部组织维护 -> 新建企业 flow.

    Args:
        driver: Existing Selenium Chrome driver that is already logged in.
        enterprise_name: 企业名称；留空时自动生成 xuyingtest企业 + 时间戳。
        enterprise_full_name: 企业全称；留空时默认使用 ``企业名称 + 全称``。
        remark: 备注信息；留空时自动生成自动化测试备注。
    """
    data = _generate_enterprise_data(
        enterprise_name=enterprise_name,
        enterprise_full_name=enterprise_full_name,
        remark=remark,
    )

    step_results: list[str] = []

    org_result = _ensure_org_structure_page(driver)
    step_results.append(f"步骤2-进入组织架构设置：{org_result}")
    if "未能" in org_result:
        return f"未能创建企业 {data.name}：" + " ".join(step_results)

    external_org_result = _open_external_org_maintenance(driver)
    step_results.append(f"步骤3-外部组织维护：{external_org_result}")
    if not external_org_result.startswith("已"):
        return f"未能创建企业 {data.name}：" + " ".join(step_results)

    new_dialog_result = _open_new_enterprise_dialog(driver)
    step_results.append(f"步骤4-新建企业：{new_dialog_result}")
    if not new_dialog_result.startswith(("已", "新建企业弹窗")):
        return f"未能创建企业 {data.name}：" + " ".join(step_results)

    fill_result = _fill_enterprise_form(driver, data)
    step_results.append(f"步骤5-填写：{fill_result}")
    if not fill_result.startswith("已填写"):
        return f"未能创建企业 {data.name}：" + " ".join(step_results)

    save_result = _save_new_enterprise_and_wait(driver, data.name)
    step_results.append(f"步骤5-保存：{save_result}")
    if save_result.startswith(("已保存", "企业")):
        verified = _search_enterprise_by_name(driver, data.name)
        verify_result = "已在外部组织维护列表验证可见" if verified else "保存后未能再次搜索验证可见"
        step_results.append(f"步骤5-验证：{verify_result}")
        return (
            f"已完成新建企业流程：企业名称={data.name}，企业全称={data.full_name}，"
            f"备注信息={data.remark}。" + " ".join(step_results)
        )

    return f"未能确认企业 {data.name} 创建成功：" + " ".join(step_results)


__all__ = [
    "create_new_enterprise",
    "create_new_enterprise_with_department",
    "create_new_enterprise_with_department_and_contact",
]
