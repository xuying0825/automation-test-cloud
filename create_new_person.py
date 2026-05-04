"""
Standalone Selenium helper for creating a new person in eTeams Org. Structure.

Usage:
    from create_new_person import create_new_person
    result = create_new_person(driver)  # auto-generates required field values

Precondition: ``driver`` is already logged in to eTeams. The helper will enter
``Org. Structure`` by reusing ``select_org_structure(driver)`` when needed, then
validate every step before moving to the next one:

1. Enter Org. Structure.
2. Select the left-side secondary menu item 组织维护.
3. Select the 人力资源 tab.
4. Click 新建人员.
5. Fill required fields plus one department in the new-person dialog.
6. Keep the filled form visible briefly, then save.
7. Keep the saved person's detail page visible briefly, then close it.
8. Search the 人力资源 page for the newly-created person and verify it appears.
"""

from __future__ import annotations

import logging
import os
import secrets
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any

# Setup logger
logger = logging.getLogger(__name__)

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
DEFAULT_TEST_PERSON_PREFIX = os.environ.get("ETEAMS_TEST_PERSON_PREFIX", "xuyingtest")
TEST_PERSON_NAME_OVERRIDE = os.environ.get("ETEAMS_TEST_PERSON_NAME", "").strip()
TEST_PERSON_ACCOUNT_OVERRIDE = os.environ.get("ETEAMS_TEST_PERSON_ACCOUNT", "").strip()
TEST_PERSON_MOBILE_OVERRIDE = os.environ.get("ETEAMS_TEST_PERSON_MOBILE", "").strip()
TEST_PERSON_EMAIL_OVERRIDE = os.environ.get("ETEAMS_TEST_PERSON_EMAIL", "").strip()
TEST_PERSON_EMPLOYEE_NO_OVERRIDE = os.environ.get(
    "ETEAMS_TEST_PERSON_EMPLOYEE_NO", ""
).strip()
BEFORE_SAVE_REVIEW_DELAY_SECONDS = int(
    os.environ.get("ETEAMS_NEW_PERSON_BEFORE_SAVE_DELAY_SECONDS", "5")
)
_LAST_SELECTED_DEPARTMENT = ""


@dataclass(frozen=True)
class PersonData:
    """Generated values used to fill new-person required fields."""

    name: str
    account: str
    mobile: str
    email: str
    employee_no: str
    password: str
    number: str = "1"


def _generate_person_data() -> PersonData:
    """Generate stable-looking, unique test values for this run."""
    logger.info("生成测试人员数据")
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    compact_suffix = f"{timestamp}{suffix}"
    prefix = DEFAULT_TEST_PERSON_PREFIX

    # A valid-looking Mainland China mobile number. Keep it deterministic in
    # shape but random enough to avoid colliding with prior test data.
    mobile_tail = f"{int(time.time() * 1000) % 10**9:09d}"
    mobile = TEST_PERSON_MOBILE_OVERRIDE or f"13{mobile_tail}"

    account = TEST_PERSON_ACCOUNT_OVERRIDE or f"{prefix}{compact_suffix}"
    person_data = PersonData(
        name=TEST_PERSON_NAME_OVERRIDE or f"{prefix}人员{timestamp}{suffix}",
        account=account,
        mobile=mobile,
        email=TEST_PERSON_EMAIL_OVERRIDE or f"{account}@example.com",
        employee_no=TEST_PERSON_EMPLOYEE_NO_OVERRIDE or f"EMP{compact_suffix}",
        password="Test@123456",
    )
    logger.info(f"生成的人员数据 - 姓名: {person_data.name}, 账号: {person_data.account}, 手机: {person_data.mobile}")
    return person_data


def _normalize_text(value: str | None) -> str:
    return " ".join(str(value or "").split())


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
    """Collect short visible status/error messages."""
    if fast:
        driver.implicitly_wait(0)

    selectors = [
        ".ui-message",
        ".ui-toast",
        ".ui-notification",
        ".weapp-passport-message",
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


def _body_text(driver: webdriver.Chrome) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def _dismiss_transient_overlays(driver: webdriver.Chrome) -> None:
    """Close transient menus/dropdowns that may cover the Org.Structure page."""
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.2)
    except Exception:
        pass


def _org_structure_page_is_open(driver: webdriver.Chrome) -> bool:
    """Return True when the current page is the Org. Structure page."""
    try:
        current_url = driver.current_url
        if ORG_STRUCTURE_PATH in current_url:
            return True

        body_text = _body_text(driver)
        return (
            "组织架构设置" in body_text
            and "组织维护" in body_text
            and any(
                marker in body_text
                for marker in (
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
    """Step 1: enter Org. Structure and verify it is open."""
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


def _is_org_maintenance_open(driver: webdriver.Chrome) -> bool:
    """Return True when the 组织维护 page/content is selected."""
    try:
        body_text = _body_text(driver)
        has_org_maintenance_content = any(
            marker in body_text
            for marker in (
                "人力资源",
                "行政组织",
                "部门名称",
                "部门全称",
                "新建部门",
                "新建人员",
                "Human Resources",
            )
        )
        return "组织维护" in body_text and has_org_maintenance_content
    except Exception:
        return False


def _click_left_org_maintenance_menu(driver: webdriver.Chrome) -> str:
    """Click the left-side secondary menu item 组织维护."""
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
                  '组织维护',
                  'Organization Maintenance',
                  'Org Maintenance',
                  'Organization'
                ];
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (!labels.includes(text)) continue;
                  const clickable = raw.closest('.ui-menu-list-item, li, a, button, [role="menuitem"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  if (rect.left > 300 || rect.top < 60) continue;
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


def _open_org_maintenance(driver: webdriver.Chrome) -> str:
    """Step 2: select 组织维护 in the left-side Org. Structure menu."""
    if not _org_structure_page_is_open(driver):
        return "未在 Org. Structure 页面，无法选择左侧「组织维护」。"

    _dismiss_transient_overlays(driver)

    if _is_org_maintenance_open(driver):
        return "已选择左侧二级菜单「组织维护」，并已验证组织维护页面打开。"

    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: "组织维护" in _body_text(drv)
        )
    except TimeoutException:
        return "未能选择组织维护：Org. Structure 页面未出现左侧「组织维护」菜单。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_left_org_maintenance_menu(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(_is_org_maintenance_open)
            return f"已点击左侧二级菜单 {last_clicked}，并已验证进入组织维护页面。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能选择组织维护：已尝试点击左侧「组织维护」"
        f"{'（最后点击文本：' + last_clicked + '）' if last_clicked else '，但未找到可点击菜单项'}。"
    )


def _human_resources_tab_is_active(driver: webdriver.Chrome) -> bool:
    """Return True when the 人力资源 tab is visibly selected/open."""
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
                const labels = ['人力资源', 'Human Resources', 'Human Resource', 'HR'];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (!labels.includes(text)) continue;
                  const tab = raw.closest('[role="tab"], .ui-tabs-tab, .ui-tab, li, button, a') || raw;
                  if (visible(tab) && isActive(tab)) return true;
                }
                const bodyText = normalize(document.body.innerText || document.body.textContent);
                return bodyText.includes('新建人员')
                  || bodyText.includes('新增人员')
                  || bodyText.includes('新建用户')
                  || bodyText.includes('新增用户')
                  || bodyText.includes('New Person')
                  || bodyText.includes('Add Person')
                  || bodyText.includes('New Employee')
                  || bodyText.includes('Add Employee');
                """
            )
        )
    except Exception:
        return False


def _click_human_resources_tab(driver: webdriver.Chrome) -> str:
    """Click the 人力资源 tab on the 组织维护 page."""
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
                const labels = ['人力资源', 'Human Resources', 'Human Resource', 'HR'];
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (!labels.includes(text)) continue;
                  const clickable = raw.closest('[role="tab"], .ui-tabs-tab, .ui-tab, li, button, a') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  // It should be in the main org-maintenance content, not in the left nav.
                  if (rect.left < 180 || rect.top < 60 || rect.top > 260) continue;
                  const className = String(clickable.className || '');
                  candidates.push({ element: clickable, text, left: rect.left, top: rect.top, className });
                }
                candidates.sort((a, b) =>
                  (b.className.includes('tab') ? 1 : 0) - (a.className.includes('tab') ? 1 : 0)
                  || a.top - b.top
                  || a.left - b.left
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


def _open_human_resources_tab(driver: webdriver.Chrome) -> str:
    """Step 3: select and verify the 人力资源 tab."""
    if not _is_org_maintenance_open(driver):
        return "未在组织维护页面，无法选择「人力资源」标签页。"

    if _human_resources_tab_is_active(driver):
        return "已选中「人力资源」标签页，并已验证标签页内容打开。"

    try:
        WebDriverWait(driver, 12, poll_frequency=0.4).until(
            lambda drv: "人力资源" in _body_text(drv)
            or "Human Resources" in _body_text(drv)
            or "Human Resource" in _body_text(drv)
        )
    except TimeoutException:
        return "未能选择人力资源：组织维护页面未出现「人力资源」标签页。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_human_resources_tab(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                _human_resources_tab_is_active
            )
            return f"已点击标签页 {last_clicked}，并已验证选中人力资源标签页。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能选中人力资源标签页："
        f"{'已点击 ' + last_clicked + ' 但未验证成功。' if last_clicked else '未找到可点击的人力资源标签页。'}"
    )


def _new_person_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the 新建人员 dialog/drawer is visible."""
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
                const dialogSelectors = [
                  '.ui-dialog-wrap-right',
                  '.ui-dialog-wrap',
                  '.ui-modal',
                  '.modal',
                  '.ant-modal',
                  '.el-dialog',
                  '[role="dialog"]'
                ].join(',');
                const titleMarkers = [
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
                for (const el of Array.from(document.querySelectorAll(dialogSelectors))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if (titleMarkers.some((marker) => text.includes(marker))) return true;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _click_new_person_button(driver: webdriver.Chrome) -> str:
    """Click the 新建人员 button on the 人力资源 tab."""
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
                const labels = [
                  '新建人员',
                  '新增人员',
                  '新建用户',
                  '新增用户',
                  '新建',
                  '新增',
                  'New Person',
                  'Add Person',
                  'New Employee',
                  'Add Employee',
                  'New User',
                  'Add User',
                  'New Staff',
                  'Add Staff',
                  'New',
                  'Add'
                ];
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
                  if (!labels.includes(text)) continue;
                  const rect = el.getBoundingClientRect();
                  // The toolbar sits directly under the blue eTeams header; on
                  // the real page the 新建人员 button's top can be ~58px.
                  if (rect.left < 180 || rect.top < 40) continue;
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


def _open_new_person_dialog(driver: webdriver.Chrome) -> str:
    """Step 4: click 新建人员 and verify the dialog opens."""
    if not _human_resources_tab_is_active(driver):
        return "未在人力资源标签页，无法点击「新建人员」。"

    if _new_person_dialog_is_open(driver):
        return "新建人员弹窗已打开，并已验证弹窗可见。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_person_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_person_dialog_is_open
            )
            return f"已点击 {last_clicked}，并已验证新建人员弹窗打开。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建人员弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到可点击的新建人员按钮。'}"
    )


def _get_visible_new_person_dialog(driver: webdriver.Chrome):
    """Return the visible new-person dialog WebElement."""
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
            const markers = [
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
            const dialogs = Array.from(document.querySelectorAll(selector))
              .filter((el) => visible(el) && markers.some((marker) => normalize(el.innerText || el.textContent).includes(marker)));
            dialogs.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left);
            return dialogs[0] || null;
            """
        )
    except Exception:
        return None


def _get_person_dialog_controls(driver: webdriver.Chrome) -> list[dict[str, Any]]:
    """Return visible fillable controls in the new-person dialog."""
    try:
        controls = driver.execute_script(
            """
            const dialog = arguments[0];
            if (!dialog) return [];
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
            const formItemFor = (el) => el.closest([
              '.ui-form-item',
              '.ant-form-item',
              '.el-form-item',
              '.form-item',
              '.formItem',
              '[class*="form-item"]',
              '[class*="FormItem"]',
              '[class*="field"]',
              '[class*="Field"]'
            ].join(',')) || el.parentElement;
            const labelInfoFor = (control, item) => {
              const id = control.getAttribute('id');
              let label = '';
              let nearbyRequired = false;
              if (id) {
                const labelEl = dialog.querySelector(`label[for="${CSS.escape(id)}"]`);
                label = normalize(labelEl?.innerText || labelEl?.textContent);
                nearbyRequired = nearbyRequired || /[＊*]/.test(label);
              }
              if (!label && item) {
                const labelEl = item.querySelector('label, .ui-form-item-label, .ant-form-item-label, .el-form-item__label, [class*="label"], [class*="Label"]');
                label = normalize(labelEl?.innerText || labelEl?.textContent);
                nearbyRequired = nearbyRequired || /[＊*]/.test(label);
              }
              if (!label) {
                label = normalize(control.getAttribute('aria-label') || control.getAttribute('placeholder') || control.getAttribute('name'));
              }

              // eTeams' new-person form is a table-like layout where labels are
              // often sibling text nodes/cells rather than ancestors of the
              // input. Find the nearest short visible text to the left on the
              // same row, e.g. "姓名 *", "工号", "分部 *".
              const rect = control.getBoundingClientRect();
              const rowLabels = [];
              for (const raw of Array.from(dialog.querySelectorAll('label, span, div, td, th'))) {
                if (!visible(raw) || raw === control || raw.contains(control) || control.contains(raw)) continue;
                const rawText = normalize(raw.innerText || raw.textContent);
                if (!rawText || rawText.length > 30) continue;
                if (/^请输入$|^选择日期$|^\\+$|^男$|^女$|^试用$|^矩阵团队$|^主身份$/.test(rawText)) continue;
                const rawRect = raw.getBoundingClientRect();
                const rawCenterY = rawRect.top + rawRect.height / 2;
                const controlCenterY = rect.top + rect.height / 2;
                const verticalDistance = Math.abs(rawCenterY - controlCenterY);
                const horizontallyBefore = rawRect.right <= rect.left + 24 && rawRect.right >= rect.left - 280;
                if (!horizontallyBefore || verticalDistance > Math.max(24, rect.height)) continue;
                rowLabels.push({
                  text: rawText,
                  required: /[＊*]/.test(rawText)
                    || Boolean(raw.querySelector('.required, .is-required, [class*="required"], [class*="Required"], [class*="require"], [class*="Require"]')),
                  score: verticalDistance + Math.max(0, rect.left - rawRect.right) / 10 - (/[＊*]/.test(rawText) ? 4 : 0),
                  left: rawRect.left
                });
              }
              rowLabels.sort((a, b) => a.score - b.score || b.left - a.left);
              if (rowLabels[0]?.text) {
                label = rowLabels[0].text;
                nearbyRequired = nearbyRequired || rowLabels[0].required;
              }

              label = label.replace(/[＊*]/g, '').replace(/[:：]$/, '').trim();
              if (label.length > 80) {
                const placeholder = normalize(control.getAttribute('placeholder'));
                label = placeholder || label.slice(0, 80);
              }
              return { label, nearbyRequired };
            };
            const itemIsRequired = (control, item, labelInfo) => {
              if (control.required || control.getAttribute('required') !== null || control.getAttribute('aria-required') === 'true') return true;
              if (labelInfo?.nearbyRequired) return true;
              if (!item) return false;
              const text = normalize(item.innerText || item.textContent);
              if (/[＊*]/.test(text)) return true;
              if (item.querySelector('.required, .is-required, [class*="required"], [class*="Required"], [class*="require"], [class*="Require"]')) return true;
              const itemClass = String(item.className || '').toLowerCase();
              return itemClass.includes('required') || itemClass.includes('require');
            };
            const likelyRequiredLabel = (text) => /姓名|人员名称|员工姓名|成员名称|用户名称|账号|帐号|登录名|用户名|手机号|手机号码|手机|邮箱|email|工号|编号|人员编号|employee no|employee id|account|login|user name|username|mobile|phone|member|staff/i.test(text);
            const kindFor = (control) => {
              const tag = control.tagName.toLowerCase();
              const type = String(control.getAttribute('type') || '').toLowerCase();
              const role = String(control.getAttribute('role') || '').toLowerCase();
              const cls = String(control.className || '').toLowerCase();
              if (tag === 'select' || role === 'combobox' || cls.includes('select')) return 'select';
              if (type === 'checkbox') return 'checkbox';
              if (type === 'radio') return 'radio';
              if (tag === 'textarea') return 'text';
              return 'text';
            };
            const ignoredTypes = new Set(['hidden', 'button', 'submit', 'reset', 'file', 'image']);
            const result = [];
            const seen = new Set();
            for (const control of Array.from(dialog.querySelectorAll('input, textarea, select, [contenteditable="true"], [role="combobox"]'))) {
              if (seen.has(control) || !visible(control) || control.disabled) continue;
              seen.add(control);
              const tag = control.tagName.toLowerCase();
              const type = String(control.getAttribute('type') || '').toLowerCase();
              if (ignoredTypes.has(type)) continue;
              const item = formItemFor(control);
              const labelInfo = labelInfoFor(control, item);
              const label = labelInfo.label;
              const placeholder = normalize(control.getAttribute('placeholder'));
              const name = normalize(control.getAttribute('name'));
              const title = normalize(control.getAttribute('title'));
              const ariaLabel = normalize(control.getAttribute('aria-label'));
              const itemText = normalize(item?.innerText || item?.textContent);
              const textForMatch = [label, placeholder, name, title, ariaLabel, itemText].join(' ');
              const selectedText = normalize(itemText || control.getAttribute('title') || control.value || control.textContent);
              const required = itemIsRequired(control, item, labelInfo);
              const likelyRequired = likelyRequiredLabel(textForMatch);
              const rect = control.getBoundingClientRect();
              result.push({
                element: control,
                label,
                placeholder,
                name,
                title,
                aria_label: ariaLabel,
                item_text: itemText.slice(0, 160),
                required,
                likely_required: likelyRequired,
                tag,
                type,
                role: String(control.getAttribute('role') || '').toLowerCase(),
                kind: kindFor(control),
                read_only: Boolean(control.readOnly),
                value: normalize(control.value || control.textContent || control.getAttribute('title') || selectedText),
                top: rect.top,
                left: rect.left
              });
            }
            result.sort((a, b) => a.top - b.top || a.left - b.left);
            return result;
            """,
            _get_visible_new_person_dialog(driver),
        )
    except Exception:
        return []

    normalized_controls: list[dict[str, Any]] = []
    for control in controls or []:
        if isinstance(control, dict) and control.get("element") is not None:
            normalized_controls.append(control)
    return normalized_controls


def _field_value_for_control(control: dict[str, Any], person: PersonData) -> str:
    """Choose a valid-looking value based on label/placeholder/type."""
    label_text = " ".join(
        str(control.get(key, ""))
        for key in ("label", "placeholder", "name", "title", "aria_label", "item_text")
    ).lower()
    control_type = str(control.get("type", "")).lower()

    if control_type == "password" or "密码" in label_text or "password" in label_text:
        return person.password
    if control_type == "email" or "邮箱" in label_text or "email" in label_text or "mail" in label_text:
        return person.email
    if control_type == "tel" or any(token in label_text for token in ("手机号", "手机号码", "手机", "电话", "mobile", "phone", "tel")):
        return person.mobile
    if control_type == "number" or any(token in label_text for token in ("排序", "序号", "数量", "number")):
        return person.number
    if control_type in ("date", "datetime-local") or any(
        token in label_text for token in ("日期", "入职", "转正", "date", "join")
    ):
        return datetime.now().strftime("%Y-%m-%d")
    if any(token in label_text for token in ("账号", "帐号", "登录名", "用户名", "account", "login", "username", "user name")):
        return person.account
    if any(token in label_text for token in ("工号", "编号", "人员编号", "employee no", "employee id", "code", "number")):
        return person.employee_no
    if any(token in label_text for token in ("姓名", "人员名称", "员工姓名", "成员名称", "用户名称", "name", "member", "staff")):
        return person.name
    return person.name


def _field_report_value(display_name: str, value: str) -> str:
    """Return a safe field=value string for status output."""
    if "密码" in display_name.lower() or "password" in display_name.lower():
        return f"{display_name}=******"
    return f"{display_name}={value}"


def _set_text_control_value(driver: webdriver.Chrome, element: Any, value: str) -> bool:
    """Fill a text-like control and verify its value."""
    try:
        _safe_click(driver, element)
        try:
            element.clear()
        except Exception:
            # Some custom inputs do not support clear(); Ctrl+A is a fallback.
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(Keys.BACKSPACE)
        element.send_keys(value)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const tag = input.tagName.toLowerCase();
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
            input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter' }));
            """,
            element,
            value,
        )
        current_value = str(
            driver.execute_script(
                "return arguments[0].value || arguments[0].textContent || '';", element
            )
            or ""
        ).strip()
        return current_value == value or bool(current_value)
    except Exception:
        return False


def _select_first_dropdown_option(driver: webdriver.Chrome, element: Any) -> str:
    """Open a custom/native select and choose the first available option."""
    try:
        tag_name = str(element.tag_name or "").lower()
        if tag_name == "select":
            selected = driver.execute_script(
                """
                const select = arguments[0];
                for (const option of Array.from(select.options || [])) {
                  const text = String(option.textContent || '').replace(/\\s+/g, ' ').trim();
                  if (!option.disabled && option.value !== '' && text && !/请选择|Select/.test(text)) {
                    select.value = option.value;
                    option.selected = true;
                    select.dispatchEvent(new Event('input', { bubbles: true }));
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    return text;
                  }
                }
                return '';
                """,
                element,
            )
            return str(selected or "").strip()

        _safe_click(driver, element)
        time.sleep(0.4)
        selected = driver.execute_script(
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
            const optionSelector = [
              '.ui-select-option',
              '.ui-dropdown-item',
              '.ui-menu-item',
              '.ui-tree-node-content',
              '.ui-tree-node-label',
              '.ui-tree-title',
              '.ant-select-item-option',
              '.ant-tree-node-content-wrapper',
              '.el-select-dropdown__item',
              '.el-tree-node__content',
              '.ztree li a',
              '[class*="tree-node"]',
              '[class*="TreeNode"]',
              '[class*="tree-title"]',
              '[class*="TreeTitle"]',
              '[role="option"]',
              '[role="treeitem"]'
            ].join(',');
            const badText = /无数据|暂无数据|没有数据|No Data|No options|请选择|请输入|Select|搜索|查询/i;
            const candidates = [];
            for (const el of Array.from(document.body.querySelectorAll(optionSelector))) {
              if (!visible(el)) continue;
              const text = normalize(el.innerText || el.textContent || el.getAttribute('title'));
              if (!text || badText.test(text)) continue;
              const className = String(el.className || '').toLowerCase();
              if (el.getAttribute('aria-disabled') === 'true' || className.includes('disabled')) continue;
              const rect = el.getBoundingClientRect();
              // Prefer popup options below the top application header.
              if (rect.top < 60) continue;
              // Avoid accidentally selecting the Org.Structure left menu if a
              // custom dropdown failed to open.
              if (rect.left < 300) continue;
              candidates.push({ element: el, text, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) => a.top - b.top || a.left - b.left);
            const target = candidates[0]?.element;
            if (!target) return '';
            const text = candidates[0].text;
            fireClick(target);
            return text;
            """
        )
        return str(selected or "").strip()
    except Exception:
        return ""


def _check_first_required_radio_or_checkbox(driver: webdriver.Chrome, element: Any) -> bool:
    try:
        checked = bool(
            driver.execute_script(
                """
                const control = arguments[0];
                if (control.checked) return true;
                control.click();
                control.dispatchEvent(new Event('input', { bubbles: true }));
                control.dispatchEvent(new Event('change', { bubbles: true }));
                return Boolean(control.checked);
                """,
                element,
            )
        )
        if checked:
            return True
        _safe_click(driver, element)
        return bool(element.is_selected())
    except Exception:
        return False


def _control_display_name(control: dict[str, Any]) -> str:
    for key in ("label", "placeholder", "name", "aria_label", "title"):
        value = _normalize_text(str(control.get(key, "")))
        if value:
            return value.replace("*", "").replace("＊", "")
    return "未命名字段"


def _control_match_text(control: dict[str, Any]) -> str:
    return " ".join(
        _normalize_text(str(control.get(key, "")))
        for key in ("label", "placeholder", "name", "title", "aria_label", "item_text")
    )


def _is_department_control(control: dict[str, Any]) -> bool:
    """Return True when a control appears to be the 部门 selector."""
    primary_text = " ".join(
        _normalize_text(str(control.get(key, "")))
        for key in ("label", "placeholder", "name", "title", "aria_label")
    ).lower()
    item_text = _normalize_text(str(control.get("item_text", ""))).lower()
    # item_text can be noisy if a UI library wraps multiple fields together, so
    # only use it when the direct label/placeholder/name metadata is absent.
    text = primary_text or item_text
    return any(
        marker in text
        for marker in (
            "部门",
            "所属部门",
            "主部门",
            "所在部门",
            "department",
            "dept",
        )
    )


def _control_has_meaningful_value(control: dict[str, Any]) -> bool:
    """Return True when a control value is not just a label/placeholder."""
    value = _normalize_text(str(control.get("value", "")))
    if not value:
        return False

    display_name = _control_display_name(control)
    placeholder = _normalize_text(str(control.get("placeholder", "")))
    empty_markers = (
        "请选择",
        "请输入",
        "please select",
        "please enter",
        "select",
        "choose",
    )
    value_lower = value.lower()
    if any(marker in value_lower for marker in empty_markers):
        return False
    if value in {display_name, placeholder, "无", "none", "null", "-"}:
        return False

    compact_value = value
    for token in (display_name, placeholder, "*", "＊", ":", "："):
        compact_value = compact_value.replace(token, "")
    return bool(compact_value.strip())


def _click_confirm_in_open_selection_popup(driver: webdriver.Chrome) -> str:
    """Click 确定/OK in a secondary selector popup, if one is visible."""
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
                const newPersonDialog = arguments[0];
                const popupSelector = [
                  '.ui-dialog-wrap',
                  '.ui-modal',
                  '.ant-modal',
                  '.el-dialog',
                  '.ui-popover',
                  '.ui-dropdown',
                  '.ui-select-dropdown',
                  '.ant-select-dropdown',
                  '.el-select-dropdown',
                  '[role="dialog"]'
                ].join(',');
                const popups = Array.from(document.body.querySelectorAll(popupSelector))
                  .filter((el) => {
                    if (!visible(el)) return false;
                    if (newPersonDialog && (el === newPersonDialog || newPersonDialog.contains(el))) return false;
                    return true;
                  });
                popups.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left);
                for (const popup of popups) {
                  const buttons = Array.from(popup.querySelectorAll('button, a, [role="button"], .ui-btn'))
                    .filter(visible)
                    .map((el) => ({
                      element: el,
                      text: normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label')),
                      rect: el.getBoundingClientRect()
                    }))
                    .filter((item) => ['确定', 'OK', 'Confirm'].includes(item.text));
                  buttons.sort((a, b) => b.rect.left - a.rect.left || b.rect.top - a.rect.top);
                  const target = buttons[0]?.element;
                  if (target) {
                    const text = buttons[0].text;
                    fireClick(target);
                    return text;
                  }
                }
                return '';
                """,
                _get_visible_new_person_dialog(driver),
            )
        ).strip()
    except Exception:
        return ""


def _select_department_option(driver: webdriver.Chrome, element: Any) -> str:
    """Select any available department candidate from the department selector."""
    selected_text = _select_first_dropdown_option(driver, element)
    if selected_text:
        confirmed = _click_confirm_in_open_selection_popup(driver)
        return f"{selected_text}{'（已确认）' if confirmed else ''}"

    # Fallback for keyboard-driven comboboxes.
    try:
        _safe_click(driver, element)
        time.sleep(0.3)
        element.send_keys(Keys.ARROW_DOWN)
        element.send_keys(Keys.ENTER)
        time.sleep(0.5)
        confirmed = _click_confirm_in_open_selection_popup(driver)
        value = _normalize_text(
            str(
                driver.execute_script(
                    """
                    const el = arguments[0];
                    const parent = el?.parentElement;
                    return el?.value
                      || el?.textContent
                      || el?.getAttribute('title')
                      || parent?.innerText
                      || '';
                    """,
                    element,
                )
                or ""
            )
        )
        if value and "请选择" not in value and "请输入" not in value:
            return f"{value}{'（已确认）' if confirmed else ''}"
    except Exception:
        pass
    return ""


def _department_row_selected_text(driver: webdriver.Chrome) -> str:
    """Return selected text in the 新建人员「部门」row, if present."""
    try:
        return str(
            driver.execute_script(
                """
                const dialog = arguments[0];
                if (!dialog) return '';
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
                const departmentItem = (() => {
                  for (const label of Array.from(dialog.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))) {
                    if (!visible(label)) continue;
                    const text = normalize(label.innerText || label.textContent || label.getAttribute('title'));
                    if (text !== '部门') continue;
                    const item = label.closest('.ui-formItem') || label.closest('.ui-form-col');
                    if (item && visible(item)) return item;
                  }
                  return null;
                })();
                if (!departmentItem) return '';

                const selected = [];
                for (const el of Array.from(departmentItem.querySelectorAll([
                  '.ui-browser-associative-selected-item',
                  '.ui-list-item',
                  '[title]',
                  'input',
                  'span',
                  'div'
                ].join(',')))) {
                  if (!visible(el)) continue;
                  const text = normalize(
                    el.getAttribute('title')
                    || el.value
                    || el.innerText
                    || el.textContent
                  );
                  if (!text || text === '部门' || text === '请选择' || text === '请输入') continue;
                  if (/^\\+$/.test(text)) continue;
                  if (text.length > 80) continue;
                  selected.push(text);
                }
                return selected[0] || '';
                """,
                _get_visible_new_person_dialog(driver),
            )
        ).strip()
    except Exception:
        return ""


def _department_row_is_present(driver: webdriver.Chrome) -> bool:
    """Return True when the 新建人员 dialog contains the 部门 row plus button."""
    try:
        return bool(
            driver.execute_script(
                """
                const dialog = arguments[0];
                if (!dialog) return false;
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
                for (const label of Array.from(dialog.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))) {
                  if (!visible(label)) continue;
                  const text = normalize(label.innerText || label.textContent || label.getAttribute('title'));
                  if (text !== '部门') continue;
                  const item = label.closest('.ui-formItem') || label.closest('.ui-form-col');
                  if (!item || !visible(item)) continue;
                  return Boolean(item.querySelector('.associative-search-icon, .ui-input-suffix, .Icon-add-to01, input[placeholder="请选择"]'));
                }
                return false;
                """,
                _get_visible_new_person_dialog(driver),
            )
        )
    except Exception:
        return False


def _click_department_plus_button(driver: webdriver.Chrome) -> str:
    """Click the + button in the 新建人员「部门」row."""
    try:
        return str(
            driver.execute_script(
                """
                const dialog = arguments[0];
                if (!dialog) return '';
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
                let item = null;
                for (const label of Array.from(dialog.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))) {
                  if (!visible(label)) continue;
                  const text = normalize(label.innerText || label.textContent || label.getAttribute('title'));
                  if (text !== '部门') continue;
                  item = label.closest('.ui-formItem') || label.closest('.ui-form-col');
                  if (item && visible(item)) break;
                }
                if (!item) return '';

                const candidates = [];
                for (const raw of Array.from(item.querySelectorAll([
                  '.associative-search-icon',
                  '.ui-input-suffix',
                  '.Icon-add-to01',
                  'svg',
                  'use',
                  'input[placeholder="请选择"]',
                  '.ui-input-wrap',
                  '.ui-browser-associative-search'
                ].join(',')))) {
                  if (!visible(raw)) continue;
                  const clickable = raw.closest('.ui-input-suffix, .associative-search-icon, .ui-input-wrap, .ui-browser-associative-search') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  const className = String(clickable.className || '');
                  const attrs = `${className} ${raw.tagName || ''}`.toLowerCase();
                  let score = 0;
                  if (/associative-search-icon|ui-input-suffix|icon-add-to01|add/.test(attrs)) score += 100;
                  if (raw.matches && raw.matches('input[placeholder="请选择"]')) score += 30;
                  score += rect.left / 1000;
                  candidates.push({ element: clickable, score, className, left: rect.left, top: rect.top });
                }
                candidates.sort((a, b) => b.score - a.score || b.left - a.left || a.top - b.top);
                const target = candidates[0]?.element;
                if (!target) return '';
                const label = candidates[0].className || '部门+按钮';
                fireClick(target);
                return normalize(label) || '部门+按钮';
                """,
                _get_visible_new_person_dialog(driver),
            )
        ).strip()
    except Exception:
        return ""


def _click_random_department_dropdown_candidate(driver: webdriver.Chrome) -> str:
    """Click a random candidate in the 部门 dropdown opened by the + button."""
    global _LAST_SELECTED_DEPARTMENT

    try:
        logger.debug("开始收集部门候选项")
        candidates = driver.execute_script(
            """
            const newPersonDialog = arguments[0];
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
            const badText = /无数据|暂无数据|没有数据|请选择|请输入|No Data|No options|Select|Search/i;
            const selector = [
              '.ui-browser-associative-dropdown-list .ui-list-item',
              '.ui-browser-associative-dropdown-list-item',
              '.ui-select-dropdown .ui-list-item',
              '.ui-dropdown .ui-list-item',
              '.ui-list-scrollview .ui-list-item',
              '[role="option"]'
            ].join(',');
            const candidates = [];
            const seen = new Set();
            for (const raw of Array.from(document.body.querySelectorAll(selector))) {
              if (!visible(raw)) continue;
              if (newPersonDialog && (raw === newPersonDialog || newPersonDialog.contains(raw))) continue;
              const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title'));
              if (!text || badText.test(text)) continue;
              const rect = raw.getBoundingClientRect();
              // The department dropdown opens next to the department +
              // control in the right-side new-person drawer. Ignore the
              // Org.Structure tree on the far left and the people table.
              if (rect.left < 650 || rect.top < 120) continue;
              const className = String(raw.className || '').toLowerCase();
              if (className.includes('disabled') || raw.getAttribute('aria-disabled') === 'true') continue;
              const key = `${text}\\u0000${Math.round(rect.left)}\\u0000${Math.round(rect.top)}`;
              if (seen.has(key)) continue;
              seen.add(key);
              candidates.push({ text, left: rect.left, top: rect.top });
            }
            candidates.sort((a, b) => a.top - b.top || a.left - b.left);
            return candidates.map((candidate) => candidate.text);
            """,
            _get_visible_new_person_dialog(driver),
        )
        candidates = [
            _normalize_text(str(candidate))
            for candidate in (candidates or [])
            if _normalize_text(str(candidate))
        ]
        if not candidates:
            logger.warning("未找到可用的部门候选项")
            return ""

        logger.info(f"找到 {len(candidates)} 个部门候选项")
        # Use Python's secrets module instead of browser Math.random, and avoid
        # repeating the same department back-to-back when there is a choice.
        eligible = [
            candidate
            for candidate in candidates
            if candidate != _LAST_SELECTED_DEPARTMENT
        ] or candidates
        selected_text = eligible[secrets.randbelow(len(eligible))]
        selected_index = candidates.index(selected_text)
        
        logger.info(f"随机选择部门: {selected_text} (索引: {selected_index}, 上次选择: {_LAST_SELECTED_DEPARTMENT})")

        clicked = str(
            driver.execute_script(
                """
                const newPersonDialog = arguments[0];
                const selectedIndex = arguments[1];
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
                const badText = /无数据|暂无数据|没有数据|请选择|请输入|No Data|No options|Select|Search/i;
                const candidates = [];
                const seen = new Set();
                const selector = [
                  '.ui-browser-associative-dropdown-list .ui-list-item',
                  '.ui-browser-associative-dropdown-list-item',
                  '.ui-select-dropdown .ui-list-item',
                  '.ui-dropdown .ui-list-item',
                  '.ui-list-scrollview .ui-list-item',
                  '[role="option"]'
                ].join(',');
                for (const raw of Array.from(document.body.querySelectorAll(selector))) {
                  if (!visible(raw)) continue;
                  if (newPersonDialog && (raw === newPersonDialog || newPersonDialog.contains(raw))) continue;
                  const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title'));
                  if (!text || badText.test(text)) continue;
                  const rect = raw.getBoundingClientRect();
                  // The department dropdown opens next to the department +
                  // control in the right-side new-person drawer. Ignore the
                  // Org.Structure tree on the far left and the people table.
                  if (rect.left < 650 || rect.top < 120) continue;
                  const className = String(raw.className || '').toLowerCase();
                  if (className.includes('disabled') || raw.getAttribute('aria-disabled') === 'true') continue;
                  const key = `${text}\\u0000${Math.round(rect.left)}\\u0000${Math.round(rect.top)}`;
                  if (seen.has(key)) continue;
                  seen.add(key);
                  candidates.push({ element: raw, text, left: rect.left, top: rect.top });
                }
                candidates.sort((a, b) => a.top - b.top || a.left - b.left);
                const target = candidates[selectedIndex]?.element;
                if (!target) return '';
                const text = candidates[selectedIndex].text;
                fireClick(target);
                return `${text}（随机候选 ${selectedIndex + 1}/${candidates.length}）`;
                """,
                _get_visible_new_person_dialog(driver),
                selected_index,
            )
        ).strip()
        if clicked:
            _LAST_SELECTED_DEPARTMENT = selected_text
            logger.info(f"成功点击部门: {clicked}")
        else:
            logger.warning("未能点击部门选项")
        return clicked
    except Exception as e:
        logger.error(f"选择部门时发生异常: {str(e)}", exc_info=True)
        return ""


def _fill_department_field(driver: webdriver.Chrome) -> str:
    """Fill the 新建人员「部门」field by clicking + and picking any dropdown item."""
    logger.info("开始填写部门字段")
    try:
        WebDriverWait(driver, 10, poll_frequency=0.3).until(_department_row_is_present)
    except TimeoutException:
        logger.warning("未找到「部门」字段或其 + 按钮")
        return "未找到「部门」字段或其 + 按钮。"

    existing = _department_row_selected_text(driver)
    if existing:
        logger.info(f"部门字段已有值: {existing}")
        return f"部门={existing}"

    logger.info("点击部门 + 按钮")
    clicked = _click_department_plus_button(driver)
    if not clicked:
        logger.warning("未能点击「部门」字段后的 + 按钮")
        return "未能点击「部门」字段后的 + 按钮。"

    selected = ""
    for attempt in range(3):
        logger.debug(f"尝试选择部门 (第 {attempt + 1}/3 次)")
        try:
            selected = WebDriverWait(driver, 8, poll_frequency=0.3).until(
                lambda drv: _click_random_department_dropdown_candidate(drv)
            )
            break
        except TimeoutException:
            logger.warning(f"第 {attempt + 1} 次尝试超时，重新点击 + 按钮")
            clicked = _click_department_plus_button(driver)
            if not clicked:
                break
            time.sleep(0.3)

    if not selected:
        logger.warning("已点击「部门」+ 按钮，但未能选择候选部门")
        return "已点击「部门」+ 按钮，但未能选择候选部门。"

    logger.info(f"成功选择部门: {selected}")
    try:
        confirmed_value = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: _department_row_selected_text(drv)
        )
        logger.info(f"确认部门字段值: {confirmed_value}")
    except TimeoutException:
        logger.warning("未能在部门字段中确认选择的值，使用选中的文本")
        confirmed_value = selected

    if confirmed_value and "随机候选" in selected and confirmed_value in selected:
        return f"部门={selected}"
    return f"部门={confirmed_value or selected}"


def _pause_before_save_for_review(seconds: int = BEFORE_SAVE_REVIEW_DELAY_SECONDS) -> str:
    """Pause before saving so the filled form can be visually inspected."""
    seconds = max(0, int(seconds))
    if seconds:
        time.sleep(seconds)
    return f"已在保存前停留 {seconds} 秒，方便查看已填写的新建人员信息。"


def _fill_required_person_fields(driver: webdriver.Chrome, person: PersonData) -> str:
    """Fill required new-person fields plus one department selector."""
    dialog = WebDriverWait(driver, 10, poll_frequency=0.3).until(
        lambda drv: _get_visible_new_person_dialog(drv)
    )
    if not dialog:
        return "未找到新建人员弹窗，无法填写必填字段。"

    try:
        controls = WebDriverWait(driver, 15, poll_frequency=0.4).until(
            lambda drv: _get_person_dialog_controls(drv) or False
        )
    except TimeoutException:
        return "新建人员弹窗中未找到可填写控件。"

    filled: list[str] = []
    failed: list[str] = []
    handled_elements: set[str] = set()

    for control in controls:
        element = control.get("element")
        if element is None:
            continue
        element_id = str(element.id)
        if element_id in handled_elements:
            continue

        # Only fill fields that the page marks as required (red * / required
        # class / aria-required). Do not fill "likely" fields such as 邮箱/手机
        # unless the page explicitly marks them as required. Department is
        # handled separately via the row's + button because it is not exposed as
        # a normal input value in this UI.
        if not bool(control.get("required")):
            continue

        display_name = _control_display_name(control)
        kind = str(control.get("kind") or "text")
        control_type = str(control.get("type") or "").lower()
        is_read_only = bool(control.get("read_only"))

        if _control_has_meaningful_value(control):
            filled.append(f"{display_name}=已存在")
            handled_elements.add(element_id)
            continue

        if kind == "checkbox" or control_type == "checkbox":
            ok = _check_first_required_radio_or_checkbox(driver, element)
            (filled if ok else failed).append(display_name)
            handled_elements.add(element_id)
            continue

        if kind == "radio" or control_type == "radio":
            ok = _check_first_required_radio_or_checkbox(driver, element)
            (filled if ok else failed).append(display_name)
            handled_elements.add(element_id)
            continue

        if kind == "select" or is_read_only:
            selected_text = _select_first_dropdown_option(driver, element)
            if selected_text:
                filled.append(f"{display_name}={selected_text}")
            else:
                # Some searchable selects are text inputs; try text only when
                # they are not readonly. This keeps required selection fields from
                # silently passing without a value.
                if not is_read_only:
                    value = _field_value_for_control(control, person)
                    ok = _set_text_control_value(driver, element, value)
                    (filled if ok else failed).append(_field_report_value(display_name, value) if ok else display_name)
                else:
                    failed.append(display_name)
            handled_elements.add(element_id)
            continue

        value = _field_value_for_control(control, person)
        ok = _set_text_control_value(driver, element, value)
        (filled if ok else failed).append(_field_report_value(display_name, value) if ok else display_name)
        handled_elements.add(element_id)

    if not filled and not failed:
        return "未能识别页面明确标记的必填字段，已停止保存，避免误填非必填字段。"

    if failed:
        return "必填字段填写失败：" + "、".join(failed) + "。"
    if not filled:
        return "未能识别或填写任何新建人员必填字段，已停止保存。"

    department_result = _fill_department_field(driver)
    if not department_result.startswith("部门="):
        return f"部门信息填写失败：{department_result}"
    filled.append(department_result)

    # Re-read detected controls and make sure required text-like controls are no
    # longer empty before moving to Save.
    empty_required: list[str] = []
    for control in _get_person_dialog_controls(driver):
        if not bool(control.get("required")):
            continue
        kind = str(control.get("kind") or "text")
        if kind in ("checkbox", "radio"):
            continue
        value = _normalize_text(str(control.get("value", "")))
        if not value and kind != "select":
            empty_required.append(_control_display_name(control))

    if empty_required:
        return "必填字段仍为空，停止保存：" + "、".join(empty_required) + "。"

    return "已填写并验证新建人员必填字段及部门信息：" + "；".join(filled) + "。"


def _click_save_new_person(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new-person dialog."""
    try:
        return str(
            driver.execute_script(
                """
                const dialog = arguments[0];
                if (!dialog) return '';
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
                  const rect = el.getBoundingClientRect();
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.x - a.x || b.y - a.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent || target.getAttribute('title'));
                fireClick(target);
                return clickedText;
                """,
                _get_visible_new_person_dialog(driver),
            )
        ).strip()
    except Exception:
        return ""


def _person_name_visible_in_body(driver: webdriver.Chrome, person_name: str) -> bool:
    try:
        return person_name in _body_text(driver)
    except Exception:
        return False


def _save_new_person_and_wait(driver: webdriver.Chrome, person: PersonData) -> str:
    """Save the dialog and wait until the new person is created/confirmed."""
    clicked = _click_save_new_person(driver)
    if not clicked:
        return "未找到新建人员弹窗中的保存按钮。"

    success_markers = (
        "保存成功",
        "新增成功",
        "创建成功",
        "操作成功",
        "success",
        "Success",
    )
    duplicate_markers = (
        "已存在",
        "重复",
        "duplicate",
        "already exists",
        "Already exists",
    )
    validation_markers = (
        "必填",
        "不能为空",
        "请选择",
        "请输入",
        "格式不正确",
        "invalid",
        "required",
    )

    deadline = time.monotonic() + 18
    last_messages: list[str] = []
    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"人员 {person.name} 已存在（页面提示：{joined_messages}），本次按已验证通过处理。"
            if any(marker in joined_messages for marker in success_markers):
                try:
                    WebDriverWait(driver, 8, poll_frequency=0.4).until(
                        lambda drv: not _new_person_dialog_is_open(drv)
                        or _person_name_visible_in_body(drv, person.name)
                    )
                except TimeoutException:
                    pass
                return f"已点击 {clicked}，页面提示保存成功：{joined_messages}。"
            if any(marker in joined_messages for marker in validation_markers):
                # Do not continue after validation failures; a prior step was not successful.
                return f"已点击 {clicked}，但页面提示校验失败：{joined_messages}。"

        if not _new_person_dialog_is_open(driver):
            if _person_name_visible_in_body(driver, person.name):
                return f"已保存人员 {person.name}，并已在页面列表中验证可见。"
            # The dialog closing after Save is treated as a successful save signal
            # only if there are no validation/error messages.
            if not last_messages:
                return f"已点击 {clicked}，新建人员弹窗已关闭，未发现错误提示。"
            return f"已点击 {clicked}，弹窗已关闭；最后页面提示：{'；'.join(last_messages)}。"

        time.sleep(0.4)

    if _person_name_visible_in_body(driver, person.name):
        return f"已保存人员 {person.name}，页面已显示该人员。"

    return (
        f"已点击 {clicked}，但未确认人员 {person.name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def _person_detail_page_is_open(driver: webdriver.Chrome, person_name: str) -> bool:
    """Return True when the saved person's detail drawer/page is visible."""
    try:
        return bool(
            driver.execute_script(
                """
                const personName = arguments[0];
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
                const selectors = [
                  '.ui-dialog-wrap-right',
                  '.ui-dialog-wrap',
                  '[role="dialog"]'
                ].join(',');
                for (const el of Array.from(document.querySelectorAll(selectors))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  const rect = el.getBoundingClientRect();
                  const isRightDrawer = rect.left >= window.innerWidth * 0.35
                    || String(el.className || '').includes('right');
                  const isNewPersonForm = text.includes('新建人员')
                    || text.includes('保存并新建');
                  const hasDetailMarkers =
                    text.includes('基本信息')
                    || text.includes('基本资料')
                    || text.includes('账号信息')
                    || text.includes('个人信息')
                    || text.includes('通讯信息')
                    || text.includes('上下级关系')
                    || text.includes('身份信息');
                  // After Save the UI shows a right-side detail drawer. Prefer
                  // matching the generated name, but also accept the right-side
                  // detail drawer when it no longer has the New Person title.
                  if (isRightDrawer && text.includes(personName) && !isNewPersonForm) return true;
                  if (isRightDrawer && hasDetailMarkers && !isNewPersonForm) return true;
                }
                return false;
                """,
                person_name,
            )
        )
    except Exception:
        return False


def _click_close_person_detail_button(
    driver: webdriver.Chrome,
    person_name: str,
) -> str:
    """Click the close button on the saved person's detail drawer/page."""
    try:
        return str(
            driver.execute_script(
                """
                const personName = arguments[0];
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

                const dialogSelectors = [
                  '.ui-dialog-wrap-right',
                  '.ui-dialog-wrap',
                  '[role="dialog"]'
                ].join(',');
                const allDialogs = Array.from(document.querySelectorAll(dialogSelectors))
                  .filter(visible);
                const dialogInfo = (el) => {
                  const text = normalize(el.innerText || el.textContent);
                  const rect = el.getBoundingClientRect();
                  const isRightDrawer = rect.left >= window.innerWidth * 0.35
                    || String(el.className || '').includes('right');
                  const isNewPersonForm = text.includes('新建人员')
                    || text.includes('保存并新建');
                  const hasDetailMarkers =
                    text.includes('基本信息')
                    || text.includes('基本资料')
                    || text.includes('账号信息')
                    || text.includes('个人信息')
                    || text.includes('通讯信息')
                    || text.includes('上下级关系')
                    || text.includes('身份信息');
                  return { text, rect, isRightDrawer, isNewPersonForm, hasDetailMarkers };
                };
                let dialogs = allDialogs.filter((el) => {
                    if (!visible(el)) return false;
                    const info = dialogInfo(el);
                    return info.isRightDrawer
                      && !info.isNewPersonForm
                      && (info.text.includes(personName) || info.hasDetailMarkers);
                  });
                // Fallback for cases where the post-save drawer content is slow
                // to hydrate: close the right-side non-New-Person drawer.
                if (!dialogs.length) {
                  dialogs = allDialogs.filter((el) => {
                    const info = dialogInfo(el);
                    return info.isRightDrawer && !info.isNewPersonForm;
                  });
                }
                dialogs.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left);
                const dialog = dialogs[0];
                if (!dialog) return '';

                const dialogRect = dialog.getBoundingClientRect();
                const candidates = [];
                const selector = [
                  'button',
                  'a',
                  '[role="button"]',
                  '.ui-btn',
                  '.ui-icon',
                  '.ui-icon-wrapper',
                  '[class*="close"]',
                  '[class*="Close"]',
                  '[class*="guanbi"]',
                  '[class*="cancel"]',
                  '[class*="Cancel"]',
                  'svg',
                  'i',
                  'span',
                  'div'
                ].join(',');

                for (const raw of Array.from(dialog.querySelectorAll(selector))) {
                  if (!visible(raw)) continue;
                  const clickable = raw.closest(
                    'button, a, [role="button"], .ui-btn, '
                    + '[class*="close"], [class*="Close"], '
                    + '[class*="guanbi"], [class*="icon"], [class*="Icon"]'
                  ) || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  const text = normalize(clickable.innerText || clickable.textContent);
                  const title = normalize(clickable.getAttribute('title'));
                  const aria = normalize(clickable.getAttribute('aria-label'));
                  const className = String(clickable.className || '');
                  const attrs = `${text} ${title} ${aria} ${className}`.toLowerCase();

                  if (/编辑|更多|保存|edit|more|save/.test(text.toLowerCase())) continue;

                  const explicitClose = /关闭|close|cancel|取消|guanbi|icon-close|icon_.*close|\\bclose\\b/.test(attrs);
                  const looksLikeTopRightClose =
                    rect.top >= dialogRect.top
                    && rect.top <= dialogRect.top + 90
                    && rect.left >= dialogRect.right - 110
                    && rect.width <= 80
                    && rect.height <= 80;

                  if (!explicitClose && !looksLikeTopRightClose) continue;

                  candidates.push({
                    element: clickable,
                    text: text || title || aria || className,
                    explicit: explicitClose ? 1 : 0,
                    left: rect.left,
                    top: rect.top,
                    area: rect.width * rect.height
                  });
                }

                candidates.sort((a, b) =>
                  b.explicit - a.explicit
                  || b.left - a.left
                  || a.top - b.top
                  || a.area - b.area
                );
                const target = candidates[0]?.element;
                if (target) {
                  const label = candidates[0].text || '详情页关闭按钮';
                  fireClick(target);
                  return normalize(label) || '详情页关闭按钮';
                }

                // Fallback: click the expected close-control coordinate in the
                // detail drawer's upper-right corner.
                const x = Math.floor(dialogRect.right - 24);
                const y = Math.floor(dialogRect.top + 28);
                const fallback = document.elementFromPoint(x, y) || dialog;
                for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                  fallback.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: x,
                    clientY: y
                  }));
                }
                if (typeof fallback.click === 'function') fallback.click();
                return '详情页右上角关闭坐标';
                """,
                person_name,
            )
        ).strip()
    except Exception:
        return ""


def _close_person_detail_page(driver: webdriver.Chrome, person: PersonData) -> str:
    """Close the detail page shown after saving a new person."""
    try:
        WebDriverWait(driver, 8, poll_frequency=0.4).until(
            lambda drv: _person_detail_page_is_open(drv, person.name)
        )
    except TimeoutException:
        return "未检测到新建人员详情页，无需关闭。"

    # Let the newly-created person's detail page remain visible briefly so the
    # UI does not flash closed immediately after save during demos/tests.
    time.sleep(2)

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_close_person_detail_button(driver, person.name)
        if not last_clicked:
            time.sleep(0.4)
            continue
        try:
            WebDriverWait(driver, 6, poll_frequency=0.3).until(
                lambda drv: not _person_detail_page_is_open(drv, person.name)
            )
            return f"已点击 {last_clicked}，关闭新建人员详情页。"
        except TimeoutException:
            time.sleep(0.4)

    return (
        "未能关闭新建人员详情页："
        f"{'已点击 ' + last_clicked + ' 但详情页仍可见。' if last_clicked else '未找到关闭按钮。'}"
    )


def _get_human_resources_search_input(driver: webdriver.Chrome):
    """Return the search input on the 人力资源 page, preferably right of 新建人员."""
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
            for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
              if (!visible(el) || insideVisibleDialog(el)) continue;
              const text = normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label'));
              if (!newPersonLabels.includes(text)) continue;
              const rect = el.getBoundingClientRect();
              if (rect.left < 180 || rect.top < 40) continue;
              newPersonButtons.push({ element: el, rect, text });
            }
            newPersonButtons.sort((a, b) => b.rect.left - a.rect.left || a.rect.top - b.rect.top);
            const buttonRect = newPersonButtons[0]?.rect || null;

            const candidates = [];
            for (const raw of Array.from(document.body.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
              if (!visible(raw) || insideVisibleDialog(raw)) continue;
              const tag = raw.tagName.toLowerCase();
              const type = String(raw.getAttribute('type') || '').toLowerCase();
              if (['hidden', 'password', 'checkbox', 'radio', 'file', 'button', 'submit', 'reset'].includes(type)) continue;
              const rect = raw.getBoundingClientRect();
              if (rect.left < 160 || rect.top < 40) continue;

              const parent = raw.parentElement;
              const grandParent = parent?.parentElement;
              const attrs = normalize([
                raw.getAttribute('placeholder'),
                raw.getAttribute('title'),
                raw.getAttribute('aria-label'),
                raw.getAttribute('data-placeholder'),
                raw.getAttribute('name'),
                raw.getAttribute('id'),
                raw.className,
                parent?.className,
                grandParent?.className,
                normalize(parent?.innerText || '').slice(0, 160),
                normalize(grandParent?.innerText || '').slice(0, 160)
              ].join(' '));

              let score = 0;
              if (buttonRect) {
                const inputCenterY = rect.top + rect.height / 2;
                const buttonCenterY = buttonRect.top + buttonRect.height / 2;
                const sameRow = Math.abs(inputCenterY - buttonCenterY) <= 80;
                if (sameRow) score += 120;
                if (rect.left >= buttonRect.right - 20) score += 100;
                else if (rect.right > buttonRect.left) score += 20;
                score -= Math.abs(rect.left - buttonRect.right) / 5;
              }
              if (/请输入.*姓名|姓名|人员|员工|工号|账号|搜索|查询|name|search|employee|staff|user/i.test(attrs)) score += 90;
              if (/请输入|please enter|enter/i.test(attrs)) score += 20;
              if (/search|搜索|查询|sousuo/i.test(attrs)) score += 30;
              if (tag === 'input') score += 15;
              if (type === 'search') score += 20;

              if (!buttonRect && score <= 0) continue;
              candidates.push({ element: raw, score, left: rect.left, top: rect.top });
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


def _describe_human_resources_search_input(driver: webdriver.Chrome, search_input: Any) -> str:
    """Return a short label/placeholder for the 人力资源 search input."""
    try:
        return str(
            driver.execute_script(
                """
                const el = arguments[0];
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                if (!el) return '';
                const parent = el.parentElement;
                const label = normalize([
                  el.getAttribute('placeholder'),
                  el.getAttribute('aria-label'),
                  el.getAttribute('title'),
                  el.getAttribute('data-placeholder'),
                  parent?.getAttribute('title'),
                  parent?.innerText
                ].join(' '));
                return label.slice(0, 80);
                """,
                search_input,
            )
        ).strip()
    except Exception:
        return ""


def _fill_human_resources_search_input(
    driver: webdriver.Chrome,
    search_input: Any,
    search_text: str,
) -> str:
    """Fill the 人力资源 search input and press Enter."""
    try:
        driver.execute_script(
            """
            const el = arguments[0];
            if (!el) return;
            el.scrollIntoView({block: 'center', inline: 'center'});
            el.focus();
            const setValue = (target, value) => {
              const tag = target.tagName.toLowerCase();
              if (tag === 'input' || tag === 'textarea') {
                const proto = tag === 'textarea'
                  ? window.HTMLTextAreaElement.prototype
                  : window.HTMLInputElement.prototype;
                const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                if (setter) setter.call(target, value);
                else target.value = value;
              } else if (target.isContentEditable) {
                target.textContent = value;
              }
              target.dispatchEvent(new Event('input', { bubbles: true }));
              target.dispatchEvent(new Event('change', { bubbles: true }));
            };
            setValue(el, '');
            """,
            search_input,
        )
        search_input.send_keys(search_text)
        search_input.send_keys(Keys.ENTER)
        time.sleep(0.2)
    except Exception:
        try:
            driver.execute_script(
                """
                const el = arguments[0];
                const value = arguments[1];
                if (!el) return;
                el.focus();
                const tag = el.tagName.toLowerCase();
                if (tag === 'input' || tag === 'textarea') {
                  const proto = tag === 'textarea'
                    ? window.HTMLTextAreaElement.prototype
                    : window.HTMLInputElement.prototype;
                  const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                  if (setter) setter.call(el, value);
                  else el.value = value;
                } else if (el.isContentEditable) {
                  el.textContent = value;
                }
                for (const type of ['input', 'change']) {
                  el.dispatchEvent(new Event(type, { bubbles: true }));
                }
                for (const type of ['keydown', 'keypress', 'keyup']) {
                  el.dispatchEvent(new KeyboardEvent(type, {
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
                search_text,
            )
        except Exception:
            return ""

    try:
        return str(
            driver.execute_script(
                """
                const el = arguments[0];
                if (!el) return '';
                return el.value || el.textContent || '';
                """,
                search_input,
            )
        ).strip()
    except Exception:
        return ""


def _click_human_resources_search_trigger(driver: webdriver.Chrome, search_input: Any) -> str:
    """Click a nearby search/query icon or button, if one exists."""
    try:
        return str(
            driver.execute_script(
                """
                const input = arguments[0];
                if (!input) return '';
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
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn, .ui-icon, i, svg, span'))) {
                  if (!visible(raw)) continue;
                  const clickable = raw.closest('button, a, [role="button"], .ui-btn, [class*="search"], [class*="Search"], [class*="sousuo"], [class*="icon"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  const sameRow = Math.abs((rect.top + rect.height / 2) - (inputRect.top + inputRect.height / 2)) <= 60;
                  const nearInput = rect.left >= inputRect.left - 20
                    && rect.left <= inputRect.right + 160
                    && sameRow;
                  if (!nearInput) continue;

                  const text = normalize(clickable.innerText || clickable.textContent);
                  const title = normalize(clickable.getAttribute('title'));
                  const aria = normalize(clickable.getAttribute('aria-label'));
                  const className = String(clickable.className || '');
                  const attrs = `${text} ${title} ${aria} ${className}`.toLowerCase();
                  if (/新建|新增|保存|关闭|取消|clear|close|cancel|add|new|save/.test(attrs)) continue;
                  const explicitSearch = /搜索|查询|search|query|sousuo|icon-search|searchicon|magnifier/.test(attrs);
                  if (!explicitSearch && rect.left < inputRect.right - 5) continue;

                  candidates.push({
                    element: clickable,
                    text: text || title || aria || className || '搜索按钮',
                    explicit: explicitSearch ? 1 : 0,
                    distance: Math.abs(rect.left - inputRect.right),
                    left: rect.left
                  });
                }
                candidates.sort((a, b) =>
                  b.explicit - a.explicit
                  || a.distance - b.distance
                  || a.left - b.left
                );
                const target = candidates[0]?.element;
                if (!target) return '';
                const label = candidates[0].text || '搜索按钮';
                fireClick(target);
                return normalize(label) || '搜索按钮';
                """,
                search_input,
            )
        ).strip()
    except Exception:
        return ""


def _human_resources_search_result_snippet(
    driver: webdriver.Chrome,
    person_name: str,
) -> str:
    """Return a visible result snippet containing the person name, excluding dialogs."""
    try:
        return str(
            driver.execute_script(
                """
                const personName = arguments[0];
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
                const selectors = [
                  'tr',
                  '[role="row"]',
                  '.ui-table-row',
                  '.ui-grid-row',
                  '.ant-table-row',
                  '.el-table__row',
                  '.table-row',
                  '.list-item',
                  'li',
                  'td',
                  '.cell',
                  'span',
                  'div'
                ].join(',');
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll(selectors))) {
                  if (!visible(el) || insideVisibleDialog(el)) continue;
                  const tag = el.tagName.toLowerCase();
                  if (['script', 'style', 'input', 'textarea'].includes(tag)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 150 || rect.top < 40) continue;
                  let text = normalize(el.innerText || el.textContent);
                  if (!text.includes(personName)) continue;
                  // Avoid treating huge page containers as the search result.
                  if (text.length > 1200) continue;
                  const hasSearchInputValue = Array.from(el.querySelectorAll('input, textarea'))
                    .some((input) => String(input.value || '').includes(personName));
                  if (hasSearchInputValue && text.length < personName.length + 20) continue;
                  candidates.push({
                    text,
                    length: text.length,
                    top: rect.top,
                    left: rect.left
                  });
                }
                candidates.sort((a, b) =>
                  a.length - b.length
                  || a.top - b.top
                  || a.left - b.left
                );
                return candidates[0]?.text.slice(0, 160) || '';
                """,
                person_name,
            )
        ).strip()
    except Exception:
        return ""


def _search_created_person_in_human_resources(
    driver: webdriver.Chrome,
    person: PersonData,
) -> str:
    """Search the 人力资源 page for the newly-created person and verify it appears."""
    if not _human_resources_tab_is_active(driver):
        return "未在人力资源标签页，无法搜索新建人员。"

    try:
        WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: not _new_person_dialog_is_open(drv)
            and not _person_detail_page_is_open(drv, person.name)
        )
    except TimeoutException:
        return "新建人员弹窗或详情页仍未关闭，无法在人员列表中搜索。"

    try:
        search_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _get_human_resources_search_input(drv)
        )
    except TimeoutException:
        return "未找到人力资源页面中位于「新建人员」按钮右侧的姓名搜索框。"

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


def _create_new_person(driver: webdriver.Chrome) -> str:
    """Run the complete new-person workflow with per-step validation."""
    logger.info("开始执行新建人员流程")
    person = _generate_person_data()
    logger.info(f"生成人员数据 - 姓名: {person.name}, 账号: {person.account}")
    step_results: list[str] = []

    logger.info("步骤1：进入 Org. Structure 页面")
    org_result = _ensure_org_structure_page(driver)
    step_results.append(f"步骤1：{org_result}")
    if not org_result.startswith(("已", "步骤", "选择")) and "已验证进入" not in org_result:
        logger.warning(f"步骤1失败: {org_result}")
        return "未能新建人员：" + " ".join(step_results)
    logger.info("步骤1完成")

    logger.info("步骤2：打开组织维护")
    org_maintenance_result = _open_org_maintenance(driver)
    step_results.append(f"步骤2：{org_maintenance_result}")
    if not org_maintenance_result.startswith("已"):
        logger.warning(f"步骤2失败: {org_maintenance_result}")
        return "未能新建人员：" + " ".join(step_results)
    logger.info("步骤2完成")

    logger.info("步骤3：打开人力资源标签页")
    hr_result = _open_human_resources_tab(driver)
    step_results.append(f"步骤3：{hr_result}")
    if not hr_result.startswith("已"):
        logger.warning(f"步骤3失败: {hr_result}")
        return "未能新建人员：" + " ".join(step_results)
    logger.info("步骤3完成")

    logger.info("步骤4：打开新建人员弹窗")
    dialog_result = _open_new_person_dialog(driver)
    step_results.append(f"步骤4：{dialog_result}")
    if not dialog_result.startswith(("已", "新建人员弹窗")):
        logger.warning(f"步骤4失败: {dialog_result}")
        return "未能新建人员：" + " ".join(step_results)
    logger.info("步骤4完成")

    logger.info("步骤5：填写人员信息")
    fill_result = _fill_required_person_fields(driver, person)
    step_results.append(f"步骤5-填写：{fill_result}")
    if not fill_result.startswith("已填写并验证"):
        logger.warning(f"步骤5-填写失败: {fill_result}")
        return "未能新建人员：" + " ".join(step_results)
    logger.info("步骤5-填写完成")

    logger.info("步骤5-保存前停留")
    review_pause_result = _pause_before_save_for_review()
    step_results.append(f"步骤5-保存前停留：{review_pause_result}")

    logger.info("步骤5-保存人员信息")
    save_result = _save_new_person_and_wait(driver, person)
    step_results.append(f"步骤5-保存：{save_result}")
    detail_open_after_save = _person_detail_page_is_open(driver, person.name)
    save_success = (
        save_result.startswith("已保存")
        or save_result.startswith("人员")
        or "页面提示保存成功" in save_result
        or "新建人员弹窗已关闭，未发现错误提示" in save_result
        or detail_open_after_save
    )
    save_failure_markers = ("校验失败", "未确认", "失败", "错误", "不能为空", "请选择", "请输入")
    
    if save_success and not any(marker in save_result for marker in save_failure_markers):
        logger.info("步骤5-保存成功，继续执行后续步骤")
        
        logger.info("步骤6：关闭人员详情页")
        close_result = _close_person_detail_page(driver, person)
        step_results.append(f"步骤6-关闭详情页：{close_result}")
        if not close_result.startswith(("已点击", "未检测到")):
            logger.warning(f"步骤6失败: {close_result}")
            return "未能完成新建人员后的详情页关闭：" + " ".join(step_results)
        logger.info("步骤6完成")

        logger.info("步骤7：搜索验证新建人员")
        search_result = _search_created_person_in_human_resources(driver, person)
        step_results.append(f"步骤7-搜索验证：{search_result}")
        if not search_result.startswith("已"):
            logger.warning(f"步骤7失败: {search_result}")
            return "未能完成新建人员后的搜索验证：" + " ".join(step_results)
        logger.info("步骤7完成")

        logger.info("新建人员流程全部完成")
        return (
            f"已完成新建人员流程，并已搜索验证创建成功。"
            f"人员姓名：{person.name}，账号：{person.account}。"
            + " ".join(step_results)
        )

    logger.warning(f"新建人员流程失败或未能确认成功")
    return "未能确认新建人员成功：" + " ".join(step_results)


def create_new_person(driver: webdriver.Chrome) -> str:
    """
    Enter Org. Structure, open 组织维护 > 人力资源, then create a new person.

    Args:
        driver: Existing Selenium Chrome driver already logged in to eTeams.

    Returns:
        Chinese status text containing each validated step and the generated
        person values that are safe to report (password is never returned).
    """
    logger.info("调用 create_new_person 工具")
    result = _create_new_person(driver)
    logger.info(f"create_new_person 结果: {result[:200]}...")
    return result


__all__ = ["create_new_person"]
