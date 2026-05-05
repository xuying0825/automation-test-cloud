"""
Standalone Selenium helper for creating a public group in eTeams Org. Structure.

Usage:
    from create_public_group import create_public_group
    result = create_public_group(driver)  # auto-generates a xuyingtest* name
    result = create_public_group(driver, group_name="my-group")

Precondition: ``driver`` is already on the Org. Structure page. Usually call
``select_org_structure(driver)`` first.
"""

import os
import secrets
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)

ORG_STRUCTURE_PATH = "/hrm/orgsetting/departmentSetting"
TEST_GROUP_NAME_OVERRIDE = os.environ.get("ETEAMS_TEST_GROUP_NAME", "").strip()
DEFAULT_TEST_GROUP_PREFIX = os.environ.get("ETEAMS_TEST_GROUP_PREFIX", "xuyingtest")
DEFAULT_IMPLICIT_WAIT_SECONDS = 5


def _generate_test_group_name() -> str:
    """Generate the public group name for this run."""
    if TEST_GROUP_NAME_OVERRIDE:
        return TEST_GROUP_NAME_OVERRIDE

    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    return f"{DEFAULT_TEST_GROUP_PREFIX}{timestamp}{suffix}"


def _safe_click(driver: webdriver.Chrome, element) -> None:
    """Click an element, falling back to JavaScript if the normal click is blocked."""
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element,
    )
    try:
        element.click()
    except (ElementClickInterceptedException, StaleElementReferenceException):
        driver.execute_script("arguments[0].click();", element)


def _collect_visible_messages(
    driver: webdriver.Chrome,
    *,
    fast: bool = False,
) -> list[str]:
    """Collect short visible status/error messages without including passwords."""
    if fast:
        driver.implicitly_wait(0)

    selectors = [
        ".ui-message",
        ".ui-toast",
        ".ui-notification",
        ".weapp-passport-message",
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


def _org_structure_page_is_open(driver: webdriver.Chrome) -> bool:
    """Return True when the current page is the Org. Structure page."""
    try:
        current_url = driver.current_url
        if ORG_STRUCTURE_PATH in current_url:
            return True

        body_text = driver.find_element(By.TAG_NAME, "body").text
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
                )
            )
        )
    except Exception:
        return False


def _is_public_group_management_open(driver: webdriver.Chrome) -> bool:
    """Return True when the Org. Structure left menu is on 群组管理."""
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        return (
            ("公共群组管理" in body_text or "Public Group" in body_text)
            and ("新建" in body_text or "New" in body_text)
        )
    except Exception:
        return False


def _new_group_dialog_is_open(driver: webdriver.Chrome) -> bool:
    """Return True if the right-side 新建群组 dialog is visible."""
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
                for (const el of Array.from(document.querySelectorAll('.ui-dialog-wrap-right, .ui-dialog-wrap'))) {
                  if (visible(el) && normalize(el.innerText || el.textContent).includes('新建群组')) {
                    return true;
                  }
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _click_left_group_management_menu(driver: webdriver.Chrome) -> str:
    """Click the left-side Org. Structure secondary menu item 群组管理."""
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
                  if (typeof el.click === 'function') {
                    el.click();
                  }
                };
                const labels = ['群组管理', 'Group Management', 'Groups Management'];
                const candidates = [];
                for (const raw of Array.from(document.body.querySelectorAll('*'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (!labels.includes(text)) continue;
                  const clickable = raw.closest('.ui-menu-list-item, li, a, button, [role="menuitem"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  if (rect.left > 280 || rect.top < 70) continue;
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


def _open_public_group_management(driver: webdriver.Chrome) -> str:
    """Open 群组管理 from the left-side Org. Structure navigation."""
    if not _org_structure_page_is_open(driver):
        return "未在 Org. Structure 页面，无法进入左侧群组管理。"

    if _is_public_group_management_open(driver):
        return "已在左侧 群组管理 / 公共群组管理 页面。"

    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass

    try:
        WebDriverWait(driver, 12).until(
            lambda drv: "群组管理" in drv.find_element(By.TAG_NAME, "body").text
        )
    except TimeoutException:
        return "未能进入群组管理：Org. Structure 页面未出现左侧「群组管理」菜单。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_left_group_management_menu(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.4).until(
                _is_public_group_management_open
            )
            return f"已点击左侧二级菜单 {last_clicked}，进入公共群组管理。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能进入群组管理：已尝试点击左侧「群组管理」"
        f"{'（最后点击文本：' + last_clicked + '）' if last_clicked else '，但未找到可点击菜单项'}。"
    )


def _click_new_group_button(driver: webdriver.Chrome) -> str:
    """Click the 新建 button on the public group management page."""
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
                  if (typeof el.click === 'function') {
                    el.click();
                  }
                };
                const candidates = [];
                for (const el of Array.from(document.body.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if (!['新建', 'New'].includes(text)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < 220 || rect.top < 40 || rect.top > 140) continue;
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.x - a.x || a.y - b.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent);
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _open_new_group_dialog(driver: webdriver.Chrome) -> str:
    """Open the 新建群组 dialog from public group management."""
    if not _is_public_group_management_open(driver):
        return "未在公共群组管理页面，无法点击新建。"

    if _new_group_dialog_is_open(driver):
        return "新建群组弹窗已打开。"

    last_clicked = ""
    for _ in range(3):
        last_clicked = _click_new_group_button(driver)
        if not last_clicked:
            time.sleep(0.5)
            continue
        try:
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                _new_group_dialog_is_open
            )
            return f"已点击 {last_clicked}，打开新建群组弹窗。"
        except TimeoutException:
            time.sleep(0.5)

    return (
        "未能打开新建群组弹窗："
        f"{'已点击 ' + last_clicked + ' 但弹窗未出现。' if last_clicked else '未找到可点击的新建按钮。'}"
    )


def _find_group_name_input(driver: webdriver.Chrome):
    """Return the 群组名称 input inside the right-side 新建群组 dialog."""
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
            const dialogs = Array.from(document.querySelectorAll('.ui-dialog-wrap-right, .ui-dialog-wrap'))
              .filter((el) => visible(el) && normalize(el.innerText || el.textContent).includes('新建群组'));
            const dialog = dialogs.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left)[0];
            if (!dialog) return null;

            const candidates = [];
            for (const input of Array.from(dialog.querySelectorAll('input'))) {
              if (!visible(input) || input.disabled || input.readOnly) continue;
              const type = String(input.getAttribute('type') || 'text').toLowerCase();
              if (!['', 'text', 'search'].includes(type)) continue;
              const className = String(input.className || '');
              if (className.includes('number')) continue;
              const rect = input.getBoundingClientRect();
              candidates.push({
                element: input,
                top: rect.top,
                left: rect.left,
                placeholder: input.getAttribute('placeholder') || '',
                value: input.value || ''
              });
            }
            candidates.sort((a, b) => a.top - b.top || a.left - b.left);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _fill_new_group_name(driver: webdriver.Chrome, group_name: str) -> str:
    """Fill the 群组名称 field in the new group dialog."""
    try:
        name_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _find_group_name_input(drv)
        )
    except TimeoutException:
        return "未找到新建群组弹窗中的「群组名称」输入框。"

    try:
        _safe_click(driver, name_input)
        name_input.clear()
        name_input.send_keys(group_name)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(
              HTMLInputElement.prototype,
              'value'
            ).set;
            setter.call(input, value);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            """,
            name_input,
            group_name,
        )
        value = str(name_input.get_attribute("value") or "").strip()
        if value == group_name:
            return f"已填写群组名称 {group_name}。"
        return f"已尝试填写群组名称，但当前输入框值为 {value or '空'}。"
    except Exception as exc:
        return f"填写群组名称失败：{str(exc)}"


def _ensure_new_group_type_is_public(driver: webdriver.Chrome) -> str:
    """Verify the new group dialog is set to 公共组."""
    try:
        dialog_text = driver.execute_script(
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
            const dialogs = Array.from(document.querySelectorAll('.ui-dialog-wrap-right, .ui-dialog-wrap'))
              .filter((el) => visible(el) && normalize(el.innerText || el.textContent).includes('新建群组'));
            const dialog = dialogs.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left)[0];
            return normalize(dialog?.innerText || dialog?.textContent || '');
            """
        )
        dialog_text = str(dialog_text or "")
        if "公共组" in dialog_text or "Public" in dialog_text:
            return "群组类型已确认是公共组。"
        return "未确认群组类型为公共组。"
    except Exception as exc:
        return f"确认群组类型失败：{str(exc)}"


def _click_save_new_group(driver: webdriver.Chrome) -> str:
    """Click 保存 in the new group dialog."""
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
                  if (typeof el.click === 'function') {
                    el.click();
                  }
                };
                const dialogs = Array.from(document.querySelectorAll('.ui-dialog-wrap-right, .ui-dialog-wrap'))
                  .filter((el) => visible(el) && normalize(el.innerText || el.textContent).includes('新建群组'));
                const dialog = dialogs.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left)[0];
                if (!dialog) return '';
                const candidates = [];
                for (const el of Array.from(dialog.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
                  if (!visible(el)) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if (!['保存', 'Save'].includes(text)) continue;
                  const rect = el.getBoundingClientRect();
                  candidates.push({ element: el, text, x: rect.left, y: rect.top });
                }
                candidates.sort((a, b) => b.x - a.x || a.y - b.y);
                const target = candidates[0]?.element;
                if (!target) return '';
                const clickedText = normalize(target.innerText || target.textContent);
                fireClick(target);
                return clickedText;
                """
            )
        ).strip()
    except Exception:
        return ""


def _group_name_visible_in_body(driver: webdriver.Chrome, group_name: str) -> bool:
    try:
        return group_name in driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return False


def _save_new_group_and_wait(driver: webdriver.Chrome, group_name: str) -> str:
    """Save the group and wait until it is created or reported as existing."""
    clicked = _click_save_new_group(driver)
    if not clicked:
        return "未找到新建群组弹窗中的保存按钮。"

    duplicate_markers = ("已存在", "重复", "duplicate", "already exists", "Already exists")
    success_markers = ("保存成功", "新增成功", "创建成功", "操作成功", "success")
    deadline = time.monotonic() + 15
    last_messages: list[str] = []
    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined_messages = "；".join(messages)
            if any(marker in joined_messages for marker in duplicate_markers):
                return f"群组 {group_name} 已存在（页面提示：{joined_messages}），视为通过。"
            if any(marker in joined_messages for marker in success_markers):
                break

        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            body_text = ""

        if any(marker in body_text for marker in duplicate_markers) and group_name in body_text:
            return f"群组 {group_name} 已存在，视为通过。"
        if _group_name_visible_in_body(driver, group_name) and not _new_group_dialog_is_open(driver):
            return f"已保存公共组 {group_name}，列表中已可见。"
        if not _new_group_dialog_is_open(driver):
            break
        time.sleep(0.4)

    if _group_name_visible_in_body(driver, group_name):
        return f"已保存公共组 {group_name}，页面已显示该群组。"

    if last_messages:
        joined_messages = "；".join(last_messages)
        if any(marker in joined_messages for marker in duplicate_markers):
            return f"群组 {group_name} 已存在（页面提示：{joined_messages}），视为通过。"

    return (
        f"已点击 {clicked}，但未在页面上确认公共组 {group_name} 保存成功"
        f"{'；页面提示：' + '；'.join(last_messages) if last_messages else ''}。"
    )


def _find_group_search_input(driver: webdriver.Chrome):
    """Return the public group list search input, outside any dialog."""
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
              const dialog = el.closest('.ui-dialog-wrap-right, .ui-dialog-wrap');
              return dialog && visible(dialog);
            };
            const candidates = [];
            for (const input of Array.from(document.querySelectorAll('input'))) {
              if (!visible(input) || insideVisibleDialog(input) || input.disabled || input.readOnly) continue;
              const placeholder = String(input.getAttribute('placeholder') || '');
              const rect = input.getBoundingClientRect();
              if (!placeholder.includes('群组名称') && !placeholder.toLowerCase().includes('group')) continue;
              candidates.push({ element: input, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) => a.top - b.top || b.left - a.left);
            return candidates[0]?.element || null;
            """
        )
    except Exception:
        return None


def _search_public_group_by_name(driver: webdriver.Chrome, group_name: str) -> bool:
    """Search the public group list by group name and return whether it is visible."""
    if _group_name_visible_in_body(driver, group_name):
        return True

    search_input = _find_group_search_input(driver)
    if not search_input:
        return _group_name_visible_in_body(driver, group_name)

    try:
        _safe_click(driver, search_input)
        search_input.clear()
        search_input.send_keys(group_name)
        search_input.send_keys(Keys.ENTER)
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(
              HTMLInputElement.prototype,
              'value'
            ).set;
            setter.call(input, value);
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
            group_name,
        )
        try:
            WebDriverWait(driver, 6, poll_frequency=0.4).until(
                lambda drv: _group_name_visible_in_body(drv, group_name)
            )
            return True
        except TimeoutException:
            return _group_name_visible_in_body(driver, group_name)
    except Exception:
        return _group_name_visible_in_body(driver, group_name)


def _create_public_group(
    driver: webdriver.Chrome,
    group_name: str = "",
) -> str:
    """
    In Org. Structure, open left 群组管理 and create/confirm a public group.

    Re-running the test is idempotent: if the group already exists, it is
    treated as a successful verification instead of deleting or duplicating it.
    """
    group_name = (group_name or "").strip() or _generate_test_group_name()

    group_management_result = _open_public_group_management(driver)
    if not group_management_result.startswith("已"):
        return f"未能创建公共组 {group_name}：{group_management_result}"

    if _search_public_group_by_name(driver, group_name):
        return (
            f"{group_management_result} 公共组 {group_name} 已存在，"
            "本次按已验证通过处理。"
        )

    new_dialog_result = _open_new_group_dialog(driver)
    if not new_dialog_result.startswith(("已", "新建群组弹窗")):
        return f"未能创建公共组 {group_name}：{group_management_result} {new_dialog_result}"

    fill_result = _fill_new_group_name(driver, group_name)
    if not fill_result.startswith("已填写"):
        return f"未能创建公共组 {group_name}：{group_management_result} {new_dialog_result} {fill_result}"

    type_result = _ensure_new_group_type_is_public(driver)
    if not type_result.startswith("群组类型已确认"):
        return f"未能创建公共组 {group_name}：{group_management_result} {new_dialog_result} {fill_result} {type_result}"

    save_result = _save_new_group_and_wait(driver, group_name)
    if save_result.startswith(("已保存", "群组")):
        verified = _search_public_group_by_name(driver, group_name)
        verify_result = "并已在公共群组管理中验证可见" if verified else "保存后未能再次搜索验证可见"
        return (
            f"{group_management_result} {new_dialog_result} {fill_result} "
            f"{type_result} {save_result} {verify_result}。"
        )

    return (
        f"未能确认公共组 {group_name} 创建成功："
        f"{group_management_result} {new_dialog_result} {fill_result} "
        f"{type_result} {save_result}"
    )


def create_public_group(driver: webdriver.Chrome, group_name: str = "") -> str:
    """
    Open left-side 群组管理 in Org. Structure and create a public group.

    Args:
        driver: Existing Selenium Chrome driver already on Org. Structure.
        group_name: Group name to create. If omitted, a random xuyingtest*
            name is generated.
    """
    return _create_public_group(driver, group_name=group_name)


__all__ = ["create_public_group"]
