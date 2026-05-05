"""
Standalone Selenium helper for selecting Org. Structure in eTeams.

Usage:
    from select_org_structure import select_org_structure
    result = select_org_structure(driver)

Precondition: ``driver`` is already logged in to eTeams. The helper performs a
real click on the top-right eTeams dropdown; it does not use a fixed URL
fallback and does not resize the browser window.
"""

import time

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

ORG_STRUCTURE_PATH = "/hrm/orgsetting/departmentSetting"


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


def _select_org_structure_from_top_right_menu(driver: webdriver.Chrome) -> str:
    """
    Open the real top-right eTeams logo dropdown and select Org. Structure.

    The dropdown is rendered only on a sufficiently wide viewport. It contains
    17 ``.e10header-dropmenu-item`` rows, with the last row being Log out
    (Chinese UI: 退出系统). Do not navigate by a hard-coded fallback URL.
    """

    def _eteams_dropdown_trigger_exists() -> bool:
        """Return whether the top-right eTeams logo/caret trigger is visible."""
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
                    const selectors = [
                      '.e10header-quick-intro',
                      '.e10header-quick-toolbar-item-logo',
                      '.e10header-quick-toolbar-item-logo-img',
                      '.e10header-quick-toolbar-item-logo-img-arrow'
                    ];
                    for (const selector of selectors) {
                      const el = document.querySelector(selector);
                      if (!visible(el)) continue;
                      const rect = el.getBoundingClientRect();
                      if (rect.top <= 80 && rect.left >= window.innerWidth - 320) {
                        return true;
                      }
                    }
                    return false;
                    """
                )
            )
        except Exception:
            return False

    def _wait_for_eteams_dropdown_trigger(timeout: int = 15) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if _eteams_dropdown_trigger_exists():
                return True
            if _org_structure_page_is_open(driver):
                return True
            time.sleep(0.5)
        return _eteams_dropdown_trigger_exists()

    def _click_eteams_dropdown() -> bool:
        """Click the top-right eTeams logo/caret dropdown using fresh DOM lookup."""
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
                    const candidates = [
                      document.querySelector('.e10header-quick-intro'),
                      document.querySelector('.e10header-quick-toolbar-item-logo'),
                      document.querySelector('.e10header-quick-toolbar-item-logo-img-arrow'),
                      document.querySelector('.e10header-quick-toolbar-item-logo-img')
                    ];
                    for (const raw of candidates) {
                      if (!visible(raw)) continue;
                      const target = raw.closest('.e10header-quick-intro') || raw;
                      if (!visible(target)) continue;
                      const rect = target.getBoundingClientRect();
                      if (rect.top <= 80 && rect.left >= window.innerWidth - 360) {
                        fireClick(target);
                        return true;
                      }
                    }
                    return false;
                    """
                )
            )
        except Exception:
            return False

    def _get_eteams_dropdown_items() -> list[str]:
        """Return visible eTeams dropdown row texts in display order."""
        try:
            items = driver.execute_script(
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
                const rows = [];
                for (const el of Array.from(document.querySelectorAll('.e10header-dropmenu-item'))) {
                  if (!visible(el)) continue;
                  const rect = el.getBoundingClientRect();
                  if (rect.left < window.innerWidth - 420 || rect.top < 40) continue;
                  const text = normalize(el.innerText || el.textContent);
                  if (text) {
                    rows.push({ text, y: rect.top, x: rect.left });
                  }
                }
                rows.sort((a, b) => a.y - b.y || a.x - b.x);
                return rows.map((row) => row.text);
                """
            )
            return [str(item).strip() for item in (items or []) if str(item).strip()]
        except Exception:
            return []

    def _wait_for_eteams_dropdown_items(timeout: int = 6) -> list[str]:
        deadline = time.monotonic() + timeout
        last_items: list[str] = []
        while time.monotonic() < deadline:
            last_items = _get_eteams_dropdown_items()
            if len(last_items) >= 17 and last_items[-1] in ("退出系统", "Log out", "Logout"):
                return last_items
            time.sleep(0.25)
        return last_items

    def _wait_for_eteams_dropdown_closed(timeout: float = 4.0) -> bool:
        """Return True after the top-right eTeams dropdown is no longer visible."""
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not _get_eteams_dropdown_items():
                return True
            time.sleep(0.2)
        return not _get_eteams_dropdown_items()

    def _dispatch_escape_and_outside_click() -> None:
        """
        Best-effort close for the eTeams dropdown.

        The menu sometimes remains mounted after selecting Org. Structure via a
        synthetic click. Dispatch Escape plus an outside click in the main page
        area so document-level close handlers run without changing data.
        """
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass

        try:
            # The eTeams menu is hover-driven (`hover-show open`). Moving the
            # real Selenium pointer away from the top-right logo is more
            # reliable than a synthetic JS click for triggering mouseleave.
            body = driver.find_element(By.TAG_NAME, "body")
            ActionChains(driver).move_to_element_with_offset(body, 0, 250).pause(
                0.2
            ).perform()
        except Exception:
            pass

        try:
            driver.execute_script(
                """
                const fire = (target, type, x, y) => {
                  target.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: x,
                    clientY: y
                  }));
                };

                document.dispatchEvent(new KeyboardEvent('keydown', {
                  bubbles: true,
                  cancelable: true,
                  key: 'Escape',
                  code: 'Escape'
                }));
                document.dispatchEvent(new KeyboardEvent('keyup', {
                  bubbles: true,
                  cancelable: true,
                  key: 'Escape',
                  code: 'Escape'
                }));

                // Click a neutral point well away from the top-right dropdown.
                const x = Math.floor(Math.min(window.innerWidth - 520, Math.max(520, window.innerWidth / 2)));
                const y = Math.floor(Math.min(window.innerHeight - 40, Math.max(140, window.innerHeight / 2)));
                const target = document.elementFromPoint(x, y) || document.body || document.documentElement;
                for (const type of ['pointerdown', 'mousedown', 'mouseup', 'click']) {
                  fire(target, type, x, y);
                }
                """
            )
        except Exception:
            pass

    def _force_hide_eteams_dropdown() -> None:
        """
        Last-resort visual cleanup for a stale eTeams dropdown.

        This does not navigate or mutate business data; it only hides the
        already-selected menu popup if the app leaves it mounted after route
        change.
        """
        try:
            driver.execute_script(
                """
                for (const popup of Array.from(document.querySelectorAll(
                  '.ui-trigger-popupInner.e10header-popup, '
                  + '.e10header-popup-content, '
                  + '.e10header-dropmenu'
                ))) {
                  if (popup.querySelector('.e10header-dropmenu-item') || popup.classList.contains('e10header-dropmenu')) {
                    popup.style.display = 'none';
                    popup.style.visibility = 'hidden';
                    popup.style.pointerEvents = 'none';
                  }
                }
                for (const trigger of Array.from(document.querySelectorAll('.e10header-quick-intro.open'))) {
                  trigger.classList.remove('open');
                }
                """
            )
        except Exception:
            pass

    def _close_eteams_dropdown_if_open(timeout: float = 5.0) -> bool:
        """Close the eTeams dropdown if it is still visible after selection."""
        if not _get_eteams_dropdown_items():
            return True

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            _dispatch_escape_and_outside_click()
            if _wait_for_eteams_dropdown_closed(timeout=0.8):
                return True

            # If Escape/outside click did not close it, the logo trigger acts as
            # a toggle. Only use it while the menu is confirmed open so we do
            # not accidentally reopen a closed dropdown.
            if _get_eteams_dropdown_items():
                _click_eteams_dropdown()
                if _wait_for_eteams_dropdown_closed(timeout=0.8):
                    return True

        _force_hide_eteams_dropdown()
        return _wait_for_eteams_dropdown_closed(timeout=1.0)

    def _click_org_structure_menu_item() -> str:
        """Click Org. Structure / 组织架构设置 in the open eTeams dropdown."""
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
                    const labels = [
                      'Org. Structure',
                      'Org.Structure',
                      'Org Structure',
                      'Organization Structure',
                      '组织架构设置',
                      '组织架构',
                      '组织结构'
                    ];
                    const candidates = [];
                    for (const el of Array.from(document.querySelectorAll('.e10header-dropmenu-item'))) {
                      if (!visible(el)) continue;
                      const rect = el.getBoundingClientRect();
                      if (rect.left < window.innerWidth - 420 || rect.top < 40) continue;
                      const text = normalize(el.innerText || el.textContent);
                      if (!text) continue;
                      const matched = labels.find((label) => text.includes(label));
                      if (!matched) continue;
                      candidates.push({ element: el, text, exact: labels.includes(text) ? 1 : 0, y: rect.top });
                    }
                    candidates.sort((a, b) => b.exact - a.exact || a.y - b.y);
                    const target = candidates[0]?.element;
                    if (!target) return '';
                    const text = normalize(target.innerText || target.textContent);
                    fireClick(target);
                    return text;
                    """
                )
            ).strip()
        except Exception:
            return ""

    try:
        time.sleep(1.0)

        if _org_structure_page_is_open(driver):
            return "已在 Org. Structure 页面，无需再次从右上角菜单选择。"

        if not _wait_for_eteams_dropdown_trigger():
            return (
                "未能选择 Org. Structure：当前窗口下未找到可见的右上角 "
                "eTeams logo 下拉菜单触发器；已按要求未调整窗口宽度。"
            )

        last_items: list[str] = []
        for _ in range(3):
            if not _click_eteams_dropdown():
                time.sleep(0.5)
                continue

            last_items = _wait_for_eteams_dropdown_items()
            if len(last_items) >= 17 and last_items[-1] in ("退出系统", "Log out", "Logout"):
                break

            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except Exception:
                pass
            time.sleep(0.5)

        if len(last_items) < 17 or last_items[-1] not in ("退出系统", "Log out", "Logout"):
            return (
                "未能选择 Org. Structure：已找到右上角 eTeams 下拉菜单触发器，"
                f"但菜单校验失败（实际 {len(last_items)} 项，最后一项："
                f"{last_items[-1] if last_items else '无'}）。"
            )

        clicked_text = _click_org_structure_menu_item()
        if not clicked_text:
            return (
                "未能选择 Org. Structure：右上角 eTeams 下拉菜单已展开且含 17 项，"
                f"但未找到 Org. Structure/组织架构设置。菜单项：{', '.join(last_items)}"
            )

        WebDriverWait(driver, 12, poll_frequency=0.5).until(_org_structure_page_is_open)
        dropdown_closed = _close_eteams_dropdown_if_open()
        return (
            "已找到右上角 eTeams 下拉菜单（17个选项，最后一项："
            f"{last_items[-1]}），并点击 {clicked_text}"
            f"{'，且已收起下拉菜单' if dropdown_closed else '，但下拉菜单仍可见'}。"
        )
    except Exception as exc:
        return f"未能选择 Org. Structure：{str(exc)}"


def select_org_structure(driver: webdriver.Chrome) -> str:
    """Find and select Org. Structure from the real top-right eTeams dropdown."""
    return _select_org_structure_from_top_right_menu(driver)


__all__ = ["select_org_structure"]
