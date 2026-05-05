"""
Standalone Selenium helper for editing a searched person in eTeams.

Usage:
    from edit_person import edit_person
    result = edit_person(driver)  # searches xuyingtest and edits one result

Precondition: ``driver`` is already logged in to eTeams. The helper enters
``Org. Structure`` when needed, opens ``组织维护`` > ``人力资源``, searches the
name box for a keyword, randomly opens one matching person detail page, clicks
``编辑``, and modifies the ``基本资料``中的``工号``.
"""

from __future__ import annotations

import random
import re
import secrets
import time
import calendar
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

from search_persion import _get_human_resources_name_search_input

DEFAULT_EDIT_PERSON_KEYWORD = "xuyingtest"


def _normalize_text(value: str | None) -> str:
    return " ".join(str(value or "").split())


def _generate_employee_no() -> str:
    """Generate a unique, valid-looking employee number for edit tests."""
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    return f"EMPEDIT{timestamp}{suffix}"


def _generate_hire_date_value() -> str:
    """Generate a random valid hire date in yyyyMMdd format."""
    now = datetime.now()
    year = random.randint(2000, min(now.year, 2025))
    month = random.randint(1, 12)
    day = random.randint(1, calendar.monthrange(year, month)[1])
    return f"{year:04d}{month:02d}{day:02d}"


def _format_hire_date_for_input(hire_date: str) -> str:
    """Format yyyyMMdd as yyyy-MM-dd for the eTeams date-picker input."""
    raw = _normalize_text(hire_date)
    match = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", raw)
    if match:
        year, month, day = match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return raw


def _hire_date_values_equal(actual: str, expected: str) -> bool:
    """Compare date values while allowing yyyyMMdd/yyyy-MM-dd/中文日期 forms."""
    actual_digits = "".join(
        ch for ch in _format_hire_date_for_input(actual) if ch.isdigit()
    )
    expected_digits = "".join(
        ch for ch in _format_hire_date_for_input(expected) if ch.isdigit()
    )
    return bool(actual_digits and expected_digits and actual_digits == expected_digits)


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


def _ensure_human_resources_page(driver: webdriver.Chrome) -> str:
    """Open Org. Structure > 组织维护 > 人力资源 when needed."""
    from create_new_person import (
        _ensure_org_structure_page,
        _human_resources_tab_is_active,
        _open_human_resources_tab,
        _open_org_maintenance,
    )

    if _human_resources_tab_is_active(driver):
        return "已在「人力资源」标签页。"

    step_results: list[str] = []
    org_result = _ensure_org_structure_page(driver)
    step_results.append(f"进入 Org. Structure：{org_result}")
    if not org_result.startswith(("已", "步骤", "选择")) and "已验证进入" not in org_result:
        return "未能进入人力资源页面：" + " ".join(step_results)

    org_maintenance_result = _open_org_maintenance(driver)
    step_results.append(f"进入组织维护：{org_maintenance_result}")
    if not org_maintenance_result.startswith("已"):
        return "未能进入人力资源页面：" + " ".join(step_results)

    hr_result = _open_human_resources_tab(driver)
    step_results.append(f"进入人力资源：{hr_result}")
    if not hr_result.startswith("已"):
        return "未能进入人力资源页面：" + " ".join(step_results)

    return "已进入「人力资源」标签页。" + " ".join(step_results)


def _search_people_by_keyword(driver: webdriver.Chrome, keyword: str) -> str:
    """Fill the 人力资源 name search box and trigger search."""
    from create_new_person import (
        _click_human_resources_search_trigger,
        _describe_human_resources_search_input,
        _fill_human_resources_search_input,
    )

    try:
        search_input = WebDriverWait(driver, 10, poll_frequency=0.3).until(
            lambda drv: _get_human_resources_name_search_input(drv)
        )
    except TimeoutException:
        return "未找到 placeholder 包含「请输入姓名」的人员搜索框。"

    input_label = _describe_human_resources_search_input(driver, search_input)
    input_value = _fill_human_resources_search_input(driver, search_input, keyword)
    if keyword not in input_value:
        return (
            f"未能确认已在人员搜索框输入 {keyword}"
            f"（当前搜索框值：{input_value or '空'}）。"
        )

    clicked_search = _click_human_resources_search_trigger(driver, search_input)
    return (
        f"已在人员搜索框{f'（{input_label}）' if input_label else ''}输入 {keyword}"
        f"{'，并点击 ' + clicked_search if clicked_search else '，并按 Enter'}。"
    )


def _get_matching_person_rows(driver: webdriver.Chrome, keyword: str) -> list[Any]:
    """Return visible person result rows containing the keyword."""
    try:
        rows = driver.execute_script(
            """
            const keyword = String(arguments[0] || '').toLowerCase();
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
            const insideOrgTree = (el) => Boolean(el.closest([
              '.ui-tree',
              '.ui-tree-warp',
              '.side-outer-left-col',
              '.weapp-hrm-com-layout-left',
              '[class*="layout-left"]'
            ].join(',')));
            const noDataPattern = /无数据|暂无数据|没有数据|没有匹配|未找到|No Data|No records|No results/i;
            const controlPattern = /新建人员|新增人员|请输入姓名|搜索|查询|New Person|Add Person|Search/i;

            const isHeaderRow = (el, text) => {
              const className = String(el.className || '').toLowerCase();
              if (el.matches('thead tr, .ant-table-thead tr, .el-table__header tr')) return true;
              if (el.querySelector('th')) return true;
              if (/header|thead|column-title/.test(className)) return true;
              const headerMarkers = ['姓名', '账号', '工号', '手机号', '部门', '岗位'];
              const headerHitCount = headerMarkers.filter((marker) => text.includes(marker)).length;
              return headerHitCount >= 3 && !text.toLowerCase().includes(keyword);
            };

            const accepted = [];
            const seen = new Set();
            const addRow = (el) => {
              if (!visible(el) || insideVisibleDialog(el) || insideOrgTree(el)) return;
              const tag = el.tagName.toLowerCase();
              if (['script', 'style', 'input', 'textarea', 'button'].includes(tag)) return;
              const rect = el.getBoundingClientRect();
              if (rect.left < 150 || rect.top < 40) return;
              const text = normalize(el.innerText || el.textContent);
              if (!text || text.length > 1200) return;
              if (!text.toLowerCase().includes(keyword)) return;
              if (noDataPattern.test(text) && text.length < 120) return;
              if (controlPattern.test(text) && text.length < 180 && !text.toLowerCase().startsWith(keyword)) return;
              if (isHeaderRow(el, text)) return;
              const ownsSearchInput = Array.from(el.querySelectorAll('input, textarea')).some((input) => {
                const placeholder = normalize(input.getAttribute('placeholder'));
                const value = normalize(input.value || input.textContent);
                return /请输入\\s*姓名|请输入姓名|姓名/.test(placeholder)
                  || value.toLowerCase() === keyword;
              });
              if (ownsSearchInput) return;
              const hasPersonNameCell = Boolean(el.querySelector(
                '.user-info-name, .weapp-hrm-com-member-list-data-list-username, [class*="member-list-data-list-username"]'
              ));
              const className = String(el.className || '').toLowerCase();
              const looksLikeTableRow =
                el.tagName.toLowerCase() === 'tr'
                || el.getAttribute('role') === 'row'
                || /table|grid|row|tr/.test(className);
              // Do not return department tree/list containers while the
              // people table is still loading. Person result rows have the
              // name/avatar column (`user-info-name`) and/or a table/grid row
              // wrapper; the left org tree can also contain xuyingtest
              // department names, but clicking it will never open a person.
              if (!hasPersonNameCell && !looksLikeTableRow) return;
              if (seen.has(el)) return;
              seen.add(el);
              accepted.push({ element: el, text, top: rect.top, left: rect.left, area: rect.width * rect.height });
            };

            const selectorGroups = [
              'tr.ui-table-grid-tr, .ui-table-grid-tr, tbody tr, .ant-table-tbody tr, .el-table__body tr',
              '[role="row"]',
              '.ui-table-row, .ui-grid-row, .ant-table-row, .el-table__row, .table-row',
              '.person-row, .employee-row, .staff-row, .user-row, .list-item, li'
            ];
            for (const selector of selectorGroups) {
              for (const el of Array.from(document.body.querySelectorAll(selector))) {
                addRow(el);
              }
              if (accepted.length) break;
            }

            if (!accepted.length) {
              for (const raw of Array.from(document.body.querySelectorAll('a, td, span, div'))) {
                if (!visible(raw) || insideVisibleDialog(raw) || insideOrgTree(raw)) continue;
                const text = normalize(raw.innerText || raw.textContent);
                if (!text || text.length > 300 || !text.toLowerCase().includes(keyword)) continue;
                const row = raw.closest(
                  'tr, [role="row"], .ui-table-row, .ui-grid-row, .ant-table-row, '
                  + '.el-table__row, .table-row, .list-item, li'
                ) || raw;
                addRow(row);
              }
            }

            accepted.sort((a, b) =>
              a.top - b.top
              || a.left - b.left
              || a.area - b.area
            );
            return accepted.map((item) => item.element);
            """,
            keyword,
        )
        return list(rows or [])
    except Exception:
        return []


def _wait_for_matching_person_rows(
    driver: webdriver.Chrome,
    keyword: str,
    timeout: int = 15,
) -> list[Any]:
    """Wait until search results containing the keyword are visible."""
    last_rows: list[Any] = []

    def _ready(drv: webdriver.Chrome):
        nonlocal last_rows
        last_rows = _get_matching_person_rows(drv, keyword)
        return last_rows or False

    try:
        return list(WebDriverWait(driver, timeout, poll_frequency=0.4).until(_ready))
    except TimeoutException:
        return last_rows


def _describe_row(driver: webdriver.Chrome, row: Any) -> str:
    """Return a short text snippet for a selected person row."""
    try:
        return str(
            driver.execute_script(
                """
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                return normalize(arguments[0]?.innerText || arguments[0]?.textContent).slice(0, 180);
                """,
                row,
            )
        ).strip()
    except Exception:
        return ""


def _get_person_name_click_info(
    driver: webdriver.Chrome,
    row: Any,
    keyword: str,
) -> dict[str, Any]:
    """
    Return the exact viewport coordinate of the visible name text in a row.

    eTeams only opens the detail drawer when the concrete name text is clicked.
    The name column contains an avatar plus the name, so clicking the row/cell or
    the avatar area is not enough. This helper targets the text-node rectangle
    containing the searched name.
    """
    try:
        info = driver.execute_script(
            """
            const row = arguments[0];
            const keyword = String(arguments[1] || '').toLowerCase();
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
            const ignored = (el) => {
              const type = String(el?.getAttribute?.('type') || '').toLowerCase();
              const className = String(el?.className || '').toLowerCase();
              return type === 'checkbox'
                || /checkbox|selection|select|avatar|photo|image|img/.test(className);
            };
            if (!row || !visible(row) || !keyword) return null;
            row.scrollIntoView({block: 'center', inline: 'nearest'});

            const rowRect = row.getBoundingClientRect();
            const candidates = [];
            const addCandidate = (element, rect, text, sourceScore = 0) => {
              if (!element || !visible(element) || !rect || rect.width <= 0 || rect.height <= 0) return;
              if (ignored(element)) return;
              const cell = element.closest('td, [role="gridcell"], .cell, [class*="cell"]') || element;
              if (!visible(cell)) return;
              const cellRect = cell.getBoundingClientRect();
              const clickable = element.closest('a, button, [role="button"]') || element;
              const clickableClass = String(clickable.className || '').toLowerCase();
              if (/checkbox|selection|select|avatar|photo|image|img/.test(clickableClass)) return;
              let score = 1000 + sourceScore;
              // The real name column is normally one of the left-most data cells.
              score -= Math.max(0, cellRect.left - rowRect.left) / 8;
              score -= Math.max(0, rect.left - cellRect.left) / 20;
              if (clickable.matches('a, button, [role="button"]')) score += 150;
              if (text.toLowerCase().startsWith(keyword)) score += 120;
              if (text.length <= 80) score += 80;
              candidates.push({
                element,
                clickable,
                text,
                score,
                left: rect.left,
                top: rect.top,
                width: rect.width,
                height: rect.height,
                cellLeft: cellRect.left
              });
            };

            // Prefer text nodes so the click lands on the concrete name text,
            // not the avatar or the full row/cell wrapper.
            const walker = document.createTreeWalker(row, NodeFilter.SHOW_TEXT, {
              acceptNode(node) {
                const text = normalize(node.nodeValue || '');
                if (!text || !text.toLowerCase().includes(keyword)) return NodeFilter.FILTER_REJECT;
                const parent = node.parentElement;
                if (!parent || !visible(parent) || ignored(parent)) return NodeFilter.FILTER_REJECT;
                return NodeFilter.FILTER_ACCEPT;
              }
            });
            let node;
            while ((node = walker.nextNode())) {
              const parent = node.parentElement;
              const text = normalize(node.nodeValue || parent.innerText || parent.textContent);
              const range = document.createRange();
              range.selectNodeContents(node);
              for (const rect of Array.from(range.getClientRects())) {
                addCandidate(parent, rect, text, 300);
              }
              range.detach();
            }

            // Fallback: smallest visible element whose own/name text contains
            // the keyword. Still click its text-side center, not the row center.
            for (const raw of Array.from(row.querySelectorAll('a, button, [role="button"], td, [role="gridcell"], span, div'))) {
              if (!visible(raw) || ignored(raw)) continue;
              const text = normalize(raw.innerText || raw.textContent);
              if (!text || text.length > 260 || !text.toLowerCase().includes(keyword)) continue;
              const rect = raw.getBoundingClientRect();
              addCandidate(raw, rect, text, raw.matches('a, button, [role="button"]') ? 180 : 0);
            }

            candidates.sort((a, b) =>
              b.score - a.score
              || a.cellLeft - b.cellLeft
              || a.left - b.left
              || a.top - b.top
            );
            const target = candidates[0];
            if (!target) return null;

            // Click inside the text rectangle. Keep the coordinate away from
            // the very left edge so it doesn't hit the avatar/name gap.
            const x = Math.floor(target.left + Math.min(Math.max(target.width * 0.45, 6), Math.max(target.width - 3, 1)));
            const y = Math.floor(target.top + target.height / 2);
            return {
              element: target.element,
              clickable: target.clickable,
              text: target.text.slice(0, 160),
              x,
              y,
              row_left: rowRect.left,
              row_top: rowRect.top
            };
            """,
            row,
            keyword,
        )
        return dict(info or {})
    except Exception:
        return {}


def _get_person_name_element(driver: webdriver.Chrome, row: Any, keyword: str):
    """Return the concrete name element in the name/avatar column."""
    try:
        return driver.execute_script(
            """
            const row = arguments[0];
            const keyword = String(arguments[1] || '').toLowerCase();
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
            if (!row || !visible(row) || !keyword) return null;
            row.scrollIntoView({block: 'center', inline: 'nearest'});
            const rowRect = row.getBoundingClientRect();
            const candidates = [];

            const add = (el, score = 0) => {
              if (!el || !visible(el)) return;
              const text = normalize(el.innerText || el.textContent || el.getAttribute('title'));
              if (!text || !text.toLowerCase().includes(keyword)) return;
              // Exclude email/phone/etc. The concrete name is rendered in the
              // avatar/name column as `.user-info-name`.
              const cell = el.closest('td, [role="gridcell"], .cell, [class*="cell"]') || el;
              if (!visible(cell)) return;
              const cellText = normalize(cell.innerText || cell.textContent);
              if (/@/.test(text) || /@/.test(cellText) && !/user-info-name/.test(String(el.className || ''))) return;
              const rect = el.getBoundingClientRect();
              const cellRect = cell.getBoundingClientRect();
              let finalScore = score;
              if (String(el.className || '').includes('user-info-name')) finalScore += 1000;
              if (el.closest('.user-info')) finalScore += 500;
              if (el.closest('.weapp-hrm-com-member-list-data-list-username')) finalScore += 300;
              if (text.toLowerCase().startsWith(keyword)) finalScore += 100;
              if (text.length <= 80) finalScore += 80;
              finalScore -= Math.max(0, cellRect.left - rowRect.left) / 6;
              candidates.push({ element: el, text, score: finalScore, left: rect.left, top: rect.top });
            };

            for (const el of Array.from(row.querySelectorAll('.user-info-name'))) add(el, 2000);
            for (const el of Array.from(row.querySelectorAll('.user-info span, .user-info div, [class*="user-info"] span'))) add(el, 1200);
            for (const el of Array.from(row.querySelectorAll('span, div, a, button, [role="button"]'))) add(el, 0);

            candidates.sort((a, b) =>
              b.score - a.score
              || a.left - b.left
              || a.top - b.top
            );
            return candidates[0]?.element || null;
            """,
            row,
            keyword,
        )
    except Exception:
        return None


def _click_person_row(driver: webdriver.Chrome, row: Any, keyword: str) -> str:
    """Click the concrete name text in a matching person row."""
    name_element = _get_person_name_element(driver, row, keyword)
    if name_element is not None:
        label = _normalize_text(
            str(
                driver.execute_script(
                    """
                    return arguments[0]?.innerText
                      || arguments[0]?.textContent
                      || arguments[0]?.getAttribute('title')
                      || '';
                    """,
                    name_element,
                )
                or "人员姓名"
            )
        )

        # First use real Selenium pointer input against the exact name element.
        # This mirrors the manual action: click the concrete name text in the
        # name column (not the avatar, row, or email cell).
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                name_element,
            )
            ActionChains(driver).move_to_element(name_element).pause(0.1).click(
                name_element
            ).perform()
            return f"{label}（具体姓名元素）"
        except Exception:
            pass

        try:
            name_element.click()
            return f"{label}（WebElement.click 具体姓名）"
        except Exception:
            pass

    info = _get_person_name_click_info(driver, row, keyword)
    if info:
        label = _normalize_text(str(info.get("text") or "人员姓名"))
        x = int(info.get("x") or 0)
        y = int(info.get("y") or 0)

        # Prefer Chrome DevTools mouse input because it clicks the exact text
        # coordinate in the viewport instead of the row/cell center.
        try:
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {"type": "mouseMoved", "x": x, "y": y, "button": "none"},
            )
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {
                    "type": "mousePressed",
                    "x": x,
                    "y": y,
                    "button": "left",
                    "clickCount": 1,
                },
            )
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {
                    "type": "mouseReleased",
                    "x": x,
                    "y": y,
                    "button": "left",
                    "clickCount": 1,
                },
            )
            return f"{label}（姓名文本坐标 {x},{y}）"
        except Exception:
            pass

        # Fallback: dispatch the click at the same text coordinate to the DOM
        # element under the point and to the located text parent/clickable.
        try:
            driver.execute_script(
                """
                const info = arguments[0];
                const elements = [
                  document.elementFromPoint(info.x, info.y),
                  info.clickable,
                  info.element
                ].filter(Boolean);
                const seen = new Set();
                const fire = (el, type) => {
                  el.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: info.x,
                    clientY: info.y
                  }));
                };
                for (const el of elements) {
                  if (seen.has(el)) continue;
                  seen.add(el);
                  for (const type of ['pointerover', 'pointermove', 'mouseover', 'mousemove', 'pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
                    fire(el, type);
                  }
                  if (typeof el.click === 'function') el.click();
                }
                """,
                info,
            )
            return f"{label}（姓名文本）"
        except Exception:
            pass

    try:
        return str(
            driver.execute_script(
                """
                const row = arguments[0];
                const keyword = String(arguments[1] || '').toLowerCase();
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
                if (!row || !visible(row)) return '';

                const candidates = [];
                for (const raw of Array.from(row.querySelectorAll('a, [role="button"], td, span, div'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  if (!text || text.length > 260 || !text.toLowerCase().includes(keyword)) continue;
                  const clickable = raw.closest('a, [role="button"]') || raw;
                  if (!visible(clickable)) continue;
                  const rect = clickable.getBoundingClientRect();
                  candidates.push({
                    element: clickable,
                    text,
                    link: clickable.matches('a, [role="button"]') ? 1 : 0,
                    length: text.length,
                    left: rect.left,
                    top: rect.top
                  });
                }

                candidates.sort((a, b) =>
                  b.link - a.link
                  || a.length - b.length
                  || a.top - b.top
                  || a.left - b.left
                );

                const target = candidates[0]?.element || row;
                const label = candidates[0]?.text || normalize(row.innerText || row.textContent);
                fireClick(target);
                return label.slice(0, 160);
                """,
                row,
                keyword,
            )
        ).strip()
    except Exception:
        try:
            _safe_click(driver, row)
            return _describe_row(driver, row)
        except Exception:
            return ""


def _open_person_detail_from_row(
    driver: webdriver.Chrome,
    row: Any,
    keyword: str,
    timeout: float = 8.0,
) -> tuple[str, Any | None]:
    """
    Open a person detail page from a result row.

    Different eTeams table renderings bind the detail opener to different
    events (single click, double click, link/cell click). Try several real and
    synthetic interactions, stopping as soon as the detail drawer appears.
    """

    def _wait_for_drawer(wait_seconds: float = 2.0):
        try:
            return WebDriverWait(driver, wait_seconds, poll_frequency=0.25).until(
                lambda drv: _get_person_detail_drawer(drv, keyword)
            )
        except TimeoutException:
            return None

    attempts: list[str] = []

    click_result = _click_person_row(driver, row, keyword)
    attempts.append(f"点击具体姓名：{click_result or '已点击'}")
    drawer = _wait_for_drawer(2.5)
    if drawer:
        return "；".join(attempts), drawer

    try:
        ActionChains(driver).move_to_element(row).double_click(row).perform()
        attempts.append("Selenium 双击行")
        drawer = _wait_for_drawer(2.5)
        if drawer:
            return "；".join(attempts), drawer
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(row).click(row).perform()
        attempts.append("Selenium 单击行")
        drawer = _wait_for_drawer(2.0)
        if drawer:
            return "；".join(attempts), drawer
    except Exception:
        pass

    try:
        strategy = str(
            driver.execute_script(
                """
                const row = arguments[0];
                const keyword = String(arguments[1] || '').toLowerCase();
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
                const fire = (el, type) => {
                  const rect = el.getBoundingClientRect();
                  const x = Math.floor(rect.left + rect.width / 2);
                  const y = Math.floor(rect.top + rect.height / 2);
                  el.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: x,
                    clientY: y,
                    detail: type === 'dblclick' ? 2 : 1
                  }));
                };
                const fireFullClick = (el) => {
                  for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    fire(el, type);
                  }
                  if (typeof el.click === 'function') el.click();
                };
                if (!row || !visible(row)) return '';

                const targets = [];
                for (const raw of Array.from(row.querySelectorAll('a, button, [role="button"], td, [role="gridcell"], span, div'))) {
                  if (!visible(raw)) continue;
                  const text = normalize(raw.innerText || raw.textContent);
                  const rect = raw.getBoundingClientRect();
                  const type = String(raw.getAttribute('type') || '').toLowerCase();
                  const className = String(raw.className || '').toLowerCase();
                  if (type === 'checkbox' || /checkbox|selection|select/.test(className)) continue;
                  let score = 0;
                  if (text.toLowerCase().includes(keyword)) score += 500;
                  if (raw.matches('a, button, [role="button"]')) score += 160;
                  if (raw.matches('td, [role="gridcell"]')) score += 80;
                  if (text && text.length <= 180) score += 40;
                  score -= Math.max(0, rect.left - row.getBoundingClientRect().left) / 30;
                  if (score <= 0) continue;
                  targets.push({ element: raw, text, score, left: rect.left, top: rect.top });
                }
                targets.sort((a, b) =>
                  b.score - a.score
                  || a.left - b.left
                  || a.top - b.top
                );

                const target = targets[0]?.element || row;
                fireFullClick(target);
                fire(target, 'dblclick');
                fireFullClick(row);
                fire(row, 'dblclick');
                return normalize(target.innerText || target.textContent || row.innerText || row.textContent).slice(0, 160);
                """,
                row,
                keyword,
            )
        ).strip()
        attempts.append(f"JS 单击/双击候选元素：{strategy or '已触发'}")
        drawer = _wait_for_drawer(max(1.0, timeout - 7.0))
        if drawer:
            return "；".join(attempts), drawer
    except Exception:
        pass

    return "；".join(attempts), None


def _get_person_detail_drawer(driver: webdriver.Chrome, keyword: str = ""):
    """Return the visible right-side person detail/edit drawer."""
    try:
        return driver.execute_script(
            """
            const keyword = String(arguments[0] || '').toLowerCase();
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
              '.ui-modal',
              '.modal',
              '.ant-modal',
              '.el-dialog',
              '[role="dialog"]'
            ].join(',');
            const detailMarkers = [
              '基本资料',
              '基本信息',
              '账号信息',
              '个人信息',
              '通讯信息',
              '上下级关系',
              '身份信息',
              '工号'
            ];
            const candidates = [];
            for (const el of Array.from(document.querySelectorAll(selectors))) {
              if (!visible(el)) continue;
              const text = normalize(el.innerText || el.textContent);
              const rect = el.getBoundingClientRect();
              const isRightDrawer = rect.left >= window.innerWidth * 0.30
                || String(el.className || '').includes('right');
              const isNewPersonForm = text.includes('新建人员') || text.includes('保存并新建');
              const hasDetailMarkers = detailMarkers.some((marker) => text.includes(marker));
              if (!isRightDrawer || isNewPersonForm || !hasDetailMarkers) continue;
              candidates.push({
                element: el,
                keywordHit: keyword && text.toLowerCase().includes(keyword) ? 1 : 0,
                right: rect.right,
                left: rect.left,
                textLength: text.length
              });
            }
            candidates.sort((a, b) =>
              b.keywordHit - a.keywordHit
              || b.right - a.right
              || b.left - a.left
              || a.textLength - b.textLength
            );
            return candidates[0]?.element || null;
            """,
            keyword,
        )
    except Exception:
        return None


def _get_edit_button_element(driver: webdriver.Chrome, drawer: Any):
    """Return the real top-right「编辑」button in the person detail drawer."""
    try:
        return driver.execute_script(
            """
            const drawer = arguments[0];
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
            if (!drawer || !visible(drawer)) return null;
            const drawerRect = drawer.getBoundingClientRect();
            const candidates = [];
            for (const raw of Array.from(drawer.querySelectorAll('button, a, [role="button"], .ui-btn'))) {
              if (!visible(raw)) continue;
              const clickable = raw.closest('button, a, [role="button"], .ui-btn') || raw;
              if (!visible(clickable)) continue;
              const text = normalize(clickable.innerText || clickable.textContent);
              const title = normalize(clickable.getAttribute('title'));
              const aria = normalize(clickable.getAttribute('aria-label'));
              const className = String(clickable.className || '');
              const attrs = `${text} ${title} ${aria} ${className}`.toLowerCase();
              if (text !== '编辑' && title !== '编辑' && aria !== '编辑' && !/\\bedit\\b|icon-edit|editicon|bianji/.test(attrs)) continue;
              if (/保存|返回|取消|关闭|删除|save|back|cancel|close|delete/.test(attrs)) continue;
              const rect = clickable.getBoundingClientRect();
              candidates.push({
                element: clickable,
                exact: text === '编辑' || title === '编辑' || aria === '编辑' ? 1 : 0,
                primary: String(clickable.className || '').includes('ui-btn-primary') ? 1 : 0,
                topArea: rect.top <= drawerRect.top + 120 ? 1 : 0,
                right: rect.right,
                top: rect.top
              });
            }
            candidates.sort((a, b) =>
              b.exact - a.exact
              || b.primary - a.primary
              || b.topArea - a.topArea
              || b.right - a.right
              || a.top - b.top
            );
            return candidates[0]?.element || null;
            """,
            drawer,
        )
    except Exception:
        return None


def _click_edit_button(driver: webdriver.Chrome, drawer: Any) -> str:
    """Click the real 编辑 button on the person detail drawer."""
    button = _get_edit_button_element(driver, drawer)
    if button is not None:
        try:
            label = _normalize_text(
                str(
                    driver.execute_script(
                        """
                        return arguments[0]?.innerText
                          || arguments[0]?.textContent
                          || arguments[0]?.getAttribute('title')
                          || arguments[0]?.getAttribute('aria-label')
                          || '编辑';
                        """,
                        button,
                    )
                    or "编辑"
                )
            )
        except Exception:
            label = "编辑"

        # Use a real Selenium pointer click first. The detail page is React-like
        # and synthetic JS clicks can report success without changing to
        #「保存/返回」edit mode.
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                button,
            )
            ActionChains(driver).move_to_element(button).pause(0.15).click(
                button
            ).perform()
            return label or "编辑"
        except Exception:
            pass

        try:
            button.click()
            return label or "编辑"
        except Exception:
            pass

        try:
            rect = driver.execute_script(
                """
                const rect = arguments[0].getBoundingClientRect();
                return {
                  x: Math.floor(rect.left + rect.width / 2),
                  y: Math.floor(rect.top + rect.height / 2)
                };
                """,
                button,
            )
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {"type": "mouseMoved", "x": rect["x"], "y": rect["y"], "button": "none"},
            )
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {
                    "type": "mousePressed",
                    "x": rect["x"],
                    "y": rect["y"],
                    "button": "left",
                    "clickCount": 1,
                },
            )
            driver.execute_cdp_cmd(
                "Input.dispatchMouseEvent",
                {
                    "type": "mouseReleased",
                    "x": rect["x"],
                    "y": rect["y"],
                    "button": "left",
                    "clickCount": 1,
                },
            )
            return label or "编辑"
        except Exception:
            pass

    # Last resort fallback for older DOM variants.
    try:
        return str(
            driver.execute_script(
                """
                const drawer = arguments[0];
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
                const target = Array.from(drawer.querySelectorAll('button, .ui-btn'))
                  .filter(visible)
                  .find((el) => normalize(el.innerText || el.textContent) === '编辑');
                if (!target) return '';
                target.click();
                return '编辑';
                """,
                drawer,
            )
        ).strip()
    except Exception:
        return ""


def _edit_mode_is_active(driver: webdriver.Chrome, drawer: Any) -> bool:
    """Return True when 编辑 has switched the drawer to 保存/返回 mode."""
    try:
        return bool(
            driver.execute_script(
                """
                const drawer = arguments[0];
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
                if (!drawer || !visible(drawer)) return false;
                const buttonTexts = Array.from(drawer.querySelectorAll('button, .ui-btn, [role="button"]'))
                  .filter(visible)
                  .map((el) => normalize(el.innerText || el.textContent || el.getAttribute('title') || el.getAttribute('aria-label')));
                const hasSave = buttonTexts.some((text) => text === '保存' || /^save$/i.test(text));
                const hasBack = buttonTexts.some((text) => text === '返回' || text === '取消' || /^back$/i.test(text) || /^cancel$/i.test(text));
                const editInputs = Array.from(drawer.querySelectorAll('input, textarea, [contenteditable="true"]'))
                  .filter(visible)
                const hasEmployeeInput = editInputs
                  .some((el) => /工号|job[_-]?num|jobnum/i.test([
                    el.getAttribute('id'),
                    el.getAttribute('name'),
                    el.getAttribute('placeholder'),
                    el.getAttribute('title'),
                    el.getAttribute('aria-label'),
                    el.closest('.ui-form-item, .ant-form-item, .el-form-item, tr, [role="row"], [class*="field"], [class*="Field"]')?.innerText
                  ].join(' ')));
                // Some renders update the editable form before the header
                // buttons' text is reflected in innerText. The job_num input is
                // the strongest signal that editing is active, because the
                // read-only detail view has no visible inputs.
                return hasEmployeeInput || ((hasSave || hasBack) && editInputs.length > 0);
                """,
                drawer,
            )
        )
    except Exception:
        return False


def _open_basic_info_section(driver: webdriver.Chrome, drawer: Any) -> None:
    """Best-effort click/scroll to the 基本资料 section before editing."""
    try:
        driver.execute_script(
            """
            const drawer = arguments[0];
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
            if (!drawer || !visible(drawer)) return '';
            const candidates = [];
            for (const raw of Array.from(drawer.querySelectorAll('button, a, [role="button"], [role="tab"], li, .ui-tabs-tab, .ui-tab, span, div'))) {
              if (!visible(raw)) continue;
              const text = normalize(raw.innerText || raw.textContent || raw.getAttribute('title') || raw.getAttribute('aria-label'));
              if (!['基本资料', '基本信息', 'Basic Info', 'Basic Information'].includes(text)) continue;
              const clickable = raw.closest('button, a, [role="button"], [role="tab"], li, .ui-tabs-tab, .ui-tab') || raw;
              const rect = clickable.getBoundingClientRect();
              candidates.push({ element: clickable, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) => a.top - b.top || a.left - b.left);
            const target = candidates[0]?.element;
            if (target) {
              target.scrollIntoView({block: 'center', inline: 'nearest'});
              fireClick(target);
            }
            return '';
            """,
            drawer,
        )
        time.sleep(0.2)
    except Exception:
        pass


def _get_employee_no_input(driver: webdriver.Chrome, drawer: Any):
    """Return the editable 工号 control in the person edit drawer."""
    try:
        return driver.execute_script(
            """
            const drawer = arguments[0];
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
            if (!drawer || !visible(drawer)) return null;
            const formItemFor = (el) => el.closest([
              '.ui-form-item',
              '.ant-form-item',
              '.el-form-item',
              '.form-item',
              '.formItem',
              '[class*="form-item"]',
              '[class*="FormItem"]',
              '[class*="field"]',
              '[class*="Field"]',
              'tr',
              '[role="row"]'
            ].join(',')) || el.parentElement;
            const labelFor = (control, item) => {
              const id = control.getAttribute('id');
              let label = '';
              if (id) {
                const labelEl = drawer.querySelector(`label[for="${CSS.escape(id)}"]`);
                label = normalize(labelEl?.innerText || labelEl?.textContent);
              }
              if (!label && item) {
                const labelEl = item.querySelector('label, .ui-form-item-label, .ant-form-item-label, .el-form-item__label, [class*="label"], [class*="Label"]');
                label = normalize(labelEl?.innerText || labelEl?.textContent);
              }
              const rect = control.getBoundingClientRect();
              const rowLabels = [];
              for (const raw of Array.from(drawer.querySelectorAll('label, span, div, td, th'))) {
                if (!visible(raw) || raw === control || raw.contains(control) || control.contains(raw)) continue;
                const rawText = normalize(raw.innerText || raw.textContent);
                if (!rawText || rawText.length > 40) continue;
                if (/^请输入$|^选择日期$|^\\+$|^男$|^女$|^试用$|^正式$/.test(rawText)) continue;
                const rawRect = raw.getBoundingClientRect();
                const rawCenterY = rawRect.top + rawRect.height / 2;
                const controlCenterY = rect.top + rect.height / 2;
                const verticalDistance = Math.abs(rawCenterY - controlCenterY);
                const horizontallyBefore = rawRect.right <= rect.left + 24 && rawRect.right >= rect.left - 320;
                if (!horizontallyBefore || verticalDistance > Math.max(28, rect.height)) continue;
                rowLabels.push({
                  text: rawText.replace(/[＊*]/g, '').replace(/[:：]$/, '').trim(),
                  score: verticalDistance + Math.max(0, rect.left - rawRect.right) / 10,
                  left: rawRect.left
                });
              }
              rowLabels.sort((a, b) => a.score - b.score || b.left - a.left);
              if (rowLabels[0]?.text) label = rowLabels[0].text;
              return label.replace(/[＊*]/g, '').replace(/[:：]$/, '').trim();
            };

            const employeePattern = /工号|人员编号|员工编号|员工号|employee\\s*(no|number|id)|employee[_-]?no|employee[_-]?id|job\\s*(no|number)|job[_-]?(num|no|number)|jobnum|work\\s*code|work[_-]?code|staff\\s*(no|id)|staff[_-]?(no|id)/i;
            const excludePattern = /账号|帐号|登录名|用户名|手机号|手机|电话|邮箱|姓名|名称|account|login|user\\s*name|username|mobile|phone|email|mail|name/i;
            const ignoredTypes = new Set(['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio']);
            const candidates = [];
            for (const control of Array.from(drawer.querySelectorAll('input, textarea, [contenteditable="true"]'))) {
              if (!visible(control) || control.disabled) continue;
              const tag = control.tagName.toLowerCase();
              const type = String(control.getAttribute('type') || '').toLowerCase();
              if (ignoredTypes.has(type)) continue;
              const item = formItemFor(control);
              const label = labelFor(control, item);
              const placeholder = normalize(control.getAttribute('placeholder'));
              const name = normalize(control.getAttribute('name'));
              const id = normalize(control.getAttribute('id'));
              const title = normalize(control.getAttribute('title'));
              const aria = normalize(control.getAttribute('aria-label'));
              const className = String(control.className || '');
              const itemText = normalize(item?.innerText || item?.textContent).slice(0, 180);
              const ownAttrs = [label, placeholder, name, id, title, aria, className].join(' ');
              const allText = [ownAttrs, itemText].join(' ');
              let score = 0;
              if (/^job[_-]?num$/i.test(id) || /^job[_-]?num$/i.test(name)) score += 1200;
              if (label === '工号') score += 800;
              if (/工号|人员编号|员工编号|员工号/.test(label)) score += 600;
              if (employeePattern.test(ownAttrs)) score += 300;
              if (employeePattern.test(allText)) score += 150;
              if (excludePattern.test(label) && !/工号|人员编号|员工编号|员工号/.test(label)) score -= 500;
              if (control.readOnly || control.getAttribute('readonly') !== null) score -= 80;
              if (tag === 'input') score += 20;
              if (score <= 0) continue;
              const rect = control.getBoundingClientRect();
              candidates.push({
                element: control,
                score,
                label,
                top: rect.top,
                left: rect.left
              });
            }
            candidates.sort((a, b) =>
              b.score - a.score
              || a.top - b.top
              || a.left - b.left
            );
            return candidates[0]?.element || null;
            """,
            drawer,
        )
    except Exception:
        return None


def _get_hire_date_input(driver: webdriver.Chrome, drawer: Any):
    """Return the editable 入职时间 date input (placeholder: 选择日期)."""
    try:
        return driver.execute_script(
            """
            const drawer = arguments[0];
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
            if (!drawer || !visible(drawer)) return null;

            const formItemFor = (el) => el.closest([
              '.ui-form-item',
              '.ant-form-item',
              '.el-form-item',
              '.form-item',
              '.formItem',
              '[class*="form-item"]',
              '[class*="FormItem"]',
              '[class*="field"]',
              '[class*="Field"]',
              'tr',
              '[role="row"]'
            ].join(',')) || el.parentElement;

            const labelFor = (control, item) => {
              const rect = control.getBoundingClientRect();
              const rowLabels = [];
              for (const raw of Array.from(drawer.querySelectorAll('label, span, div, td, th'))) {
                if (!visible(raw) || raw === control || raw.contains(control) || control.contains(raw)) continue;
                const rawText = normalize(raw.innerText || raw.textContent);
                if (!rawText || rawText.length > 40) continue;
                if (/^请选择$|^选择日期$|^请输入$|^\\+$/.test(rawText)) continue;
                const rawRect = raw.getBoundingClientRect();
                const rawCenterY = rawRect.top + rawRect.height / 2;
                const controlCenterY = rect.top + rect.height / 2;
                const verticalDistance = Math.abs(rawCenterY - controlCenterY);
                const horizontallyBefore = rawRect.right <= rect.left + 24 && rawRect.right >= rect.left - 340;
                if (!horizontallyBefore || verticalDistance > Math.max(30, rect.height)) continue;
                rowLabels.push({
                  text: rawText.replace(/[＊*]/g, '').replace(/[:：]$/, '').trim(),
                  score: verticalDistance + Math.max(0, rect.left - rawRect.right) / 10,
                  left: rawRect.left
                });
              }
              rowLabels.sort((a, b) => a.score - b.score || b.left - a.left);
              if (rowLabels[0]?.text) return rowLabels[0].text;
              return normalize(item?.innerText || item?.textContent).slice(0, 80);
            };

            const candidates = [];
            for (const input of Array.from(drawer.querySelectorAll('input, textarea'))) {
              if (!visible(input) || input.disabled) continue;
              const type = String(input.getAttribute('type') || '').toLowerCase();
              if (['hidden', 'button', 'submit', 'reset', 'file', 'image', 'checkbox', 'radio'].includes(type)) continue;
              const placeholder = normalize(input.getAttribute('placeholder'));
              const item = formItemFor(input);
              const label = labelFor(input, item);
              const attrs = [
                label,
                placeholder,
                input.getAttribute('id'),
                input.getAttribute('name'),
                input.getAttribute('title'),
                input.getAttribute('aria-label'),
                input.className,
                normalize(item?.innerText || item?.textContent).slice(0, 160)
              ].join(' ');

              let score = 0;
              if (placeholder === '选择日期') score += 500;
              if (/入职时间|入职日期|入职|join\\s*date|entry\\s*date|hire\\s*date/i.test(label)) score += 800;
              if (/入职时间|入职日期|入职|join\\s*date|entry\\s*date|hire\\s*date/i.test(attrs)) score += 300;
              if (/日期|date/i.test(attrs)) score += 80;
              if (score <= 0) continue;
              const rect = input.getBoundingClientRect();
              candidates.push({ element: input, score, top: rect.top, left: rect.left });
            }
            candidates.sort((a, b) =>
              b.score - a.score
              || a.top - b.top
              || a.left - b.left
            );
            return candidates[0]?.element || null;
            """,
            drawer,
        )
    except Exception:
        return None


def _set_input_value(driver: webdriver.Chrome, element: Any, value: str) -> str:
    """Set a text-like control value and return the final DOM value."""
    try:
        _safe_click(driver, element)
        try:
            modifier = Keys.COMMAND if hasattr(Keys, "COMMAND") else Keys.CONTROL
            element.send_keys(modifier + "a" + Keys.NULL)
            element.send_keys(Keys.BACKSPACE)
            element.send_keys(value)
            time.sleep(0.15)
            final_value = str(
                driver.execute_script(
                    "return arguments[0]?.value || arguments[0]?.textContent || '';",
                    element,
                )
                or ""
            ).strip()
            if value in final_value:
                return final_value
        except Exception:
            pass

        # Fallback only when real typing did not update the control. Prefer
        # keyboard input above because this page can leave edit mode quickly
        # after synthetic-only value changes.
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
              if (setter) setter.call(el, '');
              else el.value = '';
            } else if (el.isContentEditable) {
              el.textContent = '';
            }
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
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
            """,
            element,
            value,
        )
        time.sleep(0.2)
        final_value = str(
            driver.execute_script(
                "return arguments[0]?.value || arguments[0]?.textContent || '';",
                element,
            )
            or ""
        ).strip()
        if value not in final_value:
            try:
                element.send_keys(
                    (Keys.COMMAND if hasattr(Keys, "COMMAND") else Keys.CONTROL)
                    + "a"
                    + Keys.NULL
                )
                element.send_keys(Keys.BACKSPACE)
                element.send_keys(value)
                time.sleep(0.2)
                final_value = str(
                    driver.execute_script(
                        "return arguments[0]?.value || arguments[0]?.textContent || '';",
                        element,
                    )
                    or ""
                ).strip()
            except Exception:
                pass
        return final_value
    except Exception:
        return ""


def _click_save_button(driver: webdriver.Chrome, drawer: Any) -> str:
    """Click the 保存 button in the edit drawer."""
    try:
        return str(
            driver.execute_script(
                """
                const drawer = arguments[0];
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
                if (!drawer || !visible(drawer)) return '';
                const candidates = [];
                for (const raw of Array.from(drawer.querySelectorAll('button, a, [role="button"], .ui-btn, span, div'))) {
                  if (!visible(raw)) continue;
                  const clickable = raw.closest('button, a, [role="button"], .ui-btn') || raw;
                  if (!visible(clickable)) continue;
                  const text = normalize(clickable.innerText || clickable.textContent);
                  const title = normalize(clickable.getAttribute('title'));
                  const aria = normalize(clickable.getAttribute('aria-label'));
                  const className = String(clickable.className || '');
                  const attrs = `${text} ${title} ${aria} ${className}`.toLowerCase();
                  const saveLike = text === '保存'
                    || title === '保存'
                    || aria === '保存'
                    || /\\bsave\\b|保存/.test(attrs);
                  if (!saveLike) continue;
                  if (/保存并新建|save and/.test(attrs)) continue;
                  const rect = clickable.getBoundingClientRect();
                  candidates.push({
                    element: clickable,
                    text: text || title || aria || className || '保存',
                    exact: text === '保存' || title === '保存' || aria === '保存' ? 1 : 0,
                    top: rect.top,
                    left: rect.left
                  });
                }
                candidates.sort((a, b) =>
                  b.exact - a.exact
                  || b.left - a.left
                  || a.top - b.top
                );
                const target = candidates[0]?.element;
                if (!target) return '';
                const label = candidates[0].text || '保存';
                fireClick(target);
                return normalize(label) || '保存';
                """,
                drawer,
            )
        ).strip()
    except Exception:
        return ""


def _get_visible_calendar_day_cells(driver: webdriver.Chrome) -> list[Any]:
    """Return selectable current-month date cells from the open date dropdown."""
    try:
        cells = driver.execute_script(
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
            const panels = Array.from(document.querySelectorAll(
              '#date-picker-panel, .ui-date-picker-dayPanel, '
              + '.ui-trigger-popupInner .ui-date-picker-dayPanel, '
              + '[class*="date-picker"][class*="Panel"], '
              + '[class*="calendar"], [class*="Calendar"]'
            )).filter(visible);
            panels.sort((a, b) => b.getBoundingClientRect().top - a.getBoundingClientRect().top);
            const panel = panels.find((el) => normalize(el.innerText || el.textContent).includes('今天'))
              || panels[0];
            if (!panel) return [];

            const result = [];
            const seen = new Set();
            for (const raw of Array.from(panel.querySelectorAll('.cell, td, button, span, div'))) {
              if (!visible(raw)) continue;
              const text = normalize(raw.innerText || raw.textContent);
              if (!/^\\d{1,2}$/.test(text)) continue;
              const className = String(raw.className || '');
              const parentClass = String(raw.parentElement?.className || '');
              const attrs = `${className} ${parentClass}`.toLowerCase();
              if (/hidden|disabled|disable|outside|other|prev|next/.test(attrs)) continue;

              const clickable = raw.matches('.middle, [class*="middle"]')
                ? raw
                : (raw.querySelector('.middle, [class*="middle"]')
                  || raw.closest('.cell, td, button, a, [role="button"]')
                  || raw);
              if (!visible(clickable)) continue;
              const clickableClass = String(clickable.className || '').toLowerCase();
              if (/hidden|disabled|disable|outside|other|prev|next/.test(clickableClass)) continue;

              // eTeams marks real current-month days with
              // middle-isCurrentMonth. If the marker exists in this panel,
              // require it so we do not choose previous/next month cells.
              const panelHasCurrentMonthMarker =
                panel.querySelector('.middle-isCurrentMonth, [class*="isCurrentMonth"]');
              if (panelHasCurrentMonthMarker) {
                const marker = raw.matches('.middle-isCurrentMonth, [class*="isCurrentMonth"]')
                  || raw.querySelector('.middle-isCurrentMonth, [class*="isCurrentMonth"]')
                  || clickable.matches('.middle-isCurrentMonth, [class*="isCurrentMonth"]')
                  || clickable.querySelector('.middle-isCurrentMonth, [class*="isCurrentMonth"]');
                if (!marker) continue;
              }

              if (seen.has(clickable)) continue;
              seen.add(clickable);
              result.push(clickable);
            }
            return result;
            """
        )
        return list(cells or [])
    except Exception:
        return []


def _describe_date_cell(driver: webdriver.Chrome, cell: Any) -> str:
    """Return visible day text for a calendar cell."""
    try:
        return str(
            driver.execute_script(
                """
                return String(arguments[0]?.innerText || arguments[0]?.textContent || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                """,
                cell,
            )
        ).strip()
    except Exception:
        return ""


def _click_calendar_cell(driver: webdriver.Chrome, cell: Any) -> None:
    """Click a date cell using real pointer behavior with JS/CDP fallbacks."""
    try:
        ActionChains(driver).move_to_element(cell).pause(0.1).click(cell).perform()
        return
    except Exception:
        pass

    try:
        cell.click()
        return
    except Exception:
        pass

    try:
        rect = driver.execute_script(
            """
            const rect = arguments[0].getBoundingClientRect();
            return {
              x: Math.floor(rect.left + rect.width / 2),
              y: Math.floor(rect.top + rect.height / 2)
            };
            """,
            cell,
        )
        driver.execute_cdp_cmd(
            "Input.dispatchMouseEvent",
            {"type": "mouseMoved", "x": rect["x"], "y": rect["y"], "button": "none"},
        )
        driver.execute_cdp_cmd(
            "Input.dispatchMouseEvent",
            {
                "type": "mousePressed",
                "x": rect["x"],
                "y": rect["y"],
                "button": "left",
                "clickCount": 1,
            },
        )
        driver.execute_cdp_cmd(
            "Input.dispatchMouseEvent",
            {
                "type": "mouseReleased",
                "x": rect["x"],
                "y": rect["y"],
                "button": "left",
                "clickCount": 1,
            },
        )
        return
    except Exception:
        pass

    _safe_click(driver, cell)


def _read_visible_hire_date_value(driver: webdriver.Chrome) -> str:
    """Return the value of the visible 入职时间 date input, if any."""
    try:
        return str(
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
                const inputs = Array.from(document.querySelectorAll('input, textarea'))
                  .filter((el) => visible(el) && String(el.getAttribute('placeholder') || '').includes('选择日期'));
                const valued = inputs.find((el) => String(el.value || '').trim());
                return (valued || inputs[0])?.value || '';
                """
            )
            or ""
        ).strip()
    except Exception:
        return ""


def _pick_random_hire_date(
    driver: webdriver.Chrome,
    date_input: Any,
) -> tuple[str, str]:
    """
    Open the 入职时间 date dropdown, choose a random current-month day.

    Returns:
        (final input value, selected day text)
    """
    _safe_click(driver, date_input)
    try:
        date_input.click()
    except Exception:
        pass

    try:
        cells = WebDriverWait(driver, 6, poll_frequency=0.2).until(
            lambda drv: _get_visible_calendar_day_cells(drv)
        )
    except TimeoutException:
        return "", ""

    selected_cell = random.choice(cells)
    selected_day = _describe_date_cell(driver, selected_cell)
    _click_calendar_cell(driver, selected_cell)

    try:
        final_value = WebDriverWait(driver, 6, poll_frequency=0.2).until(
            lambda drv: _read_visible_hire_date_value(drv)
            or str(
                drv.execute_script(
                    "return arguments[0]?.value || arguments[0]?.textContent || '';",
                    date_input,
                )
                or ""
            ).strip()
        )
        return str(final_value).strip(), selected_day
    except TimeoutException:
        # Try one more date with a direct DOM click in case the first action hit
        # the cell wrapper but not the inner day text.
        retry_cells = _get_visible_calendar_day_cells(driver)
        if retry_cells:
            selected_cell = random.choice(retry_cells)
            selected_day = _describe_date_cell(driver, selected_cell)
            try:
                driver.execute_script(
                    """
                    const el = arguments[0];
                    const rect = el.getBoundingClientRect();
                    const x = Math.floor(rect.left + rect.width / 2);
                    const y = Math.floor(rect.top + rect.height / 2);
                    for (const type of ['pointerover', 'pointermove', 'mouseover', 'mousemove', 'pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
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
                    selected_cell,
                )
                final_value = WebDriverWait(driver, 4, poll_frequency=0.2).until(
                    lambda drv: _read_visible_hire_date_value(drv)
                )
                return str(final_value).strip(), selected_day
            except Exception:
                pass

        return _read_visible_hire_date_value(driver), selected_day


def _type_hire_date_value(
    driver: webdriver.Chrome,
    date_input: Any,
    hire_date: str,
) -> str:
    """Click the 入职时间 input, type yyyyMMdd, and return the typed value."""
    typed_date = _format_hire_date_for_input(hire_date)
    try:
        _safe_click(driver, date_input)
        try:
            modifier = Keys.COMMAND if hasattr(Keys, "COMMAND") else Keys.CONTROL
            date_input.send_keys(modifier + "a" + Keys.NULL)
            date_input.send_keys(Keys.BACKSPACE)
        except Exception:
            pass

        date_input.send_keys(typed_date)
        # Press Tab instead of Enter. For rows that already have an 入职时间,
        # Enter can make this page leave edit mode before Selenium clicks
        #「保存」. Tab blurs the date picker, lets it commit the value to form
        # state, and keeps the normal top-right「保存」button available.
        try:
            date_input.send_keys(Keys.TAB)
        except Exception:
            pass
        time.sleep(0.3)

        try:
            return str(
                WebDriverWait(driver, 5, poll_frequency=0.2).until(
                    lambda drv: (
                        value
                        if _hire_date_values_equal(
                            value := (
                                _read_visible_hire_date_value(drv)
                                or str(
                                    drv.execute_script(
                                        "return arguments[0]?.value || arguments[0]?.textContent || '';",
                                        date_input,
                                    )
                                    or ""
                                ).strip()
                            ),
                            typed_date,
                        )
                        else False
                    )
                )
            ).strip()
        except TimeoutException:
            pass

        return str(
            driver.execute_script(
                "return arguments[0]?.value || arguments[0]?.textContent || '';",
                date_input,
            )
            or ""
        ).strip()
    except Exception:
        try:
            driver.execute_script(
                """
                const el = arguments[0];
                const value = arguments[1];
                if (!el) return;
                el.focus();
                const setter = Object.getOwnPropertyDescriptor(
                  HTMLInputElement.prototype,
                  'value'
                )?.set;
                if (setter) setter.call(el, value);
                else el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
                """,
                date_input,
                typed_date,
            )
            return str(
                driver.execute_script(
                    "return arguments[0]?.value || arguments[0]?.textContent || '';",
                    date_input,
                )
                or ""
            ).strip()
        except Exception:
            return ""


def _wait_for_save_result(
    driver: webdriver.Chrome,
    employee_no: str,
    timeout: int = 12,
) -> str:
    """Wait for edit save feedback and return a status sentence."""
    from create_new_person import _collect_visible_messages

    success_markers = ("保存成功", "修改成功", "更新成功", "操作成功", "success")
    failure_markers = (
        "校验失败",
        "失败",
        "错误",
        "不能为空",
        "请选择",
        "请输入",
        "已存在",
        "重复",
    )
    deadline = time.monotonic() + timeout
    last_messages: list[str] = []
    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined = "；".join(messages)
            if any(marker in joined for marker in success_markers):
                return f"页面提示保存成功：{joined}"
            if any(marker in joined for marker in failure_markers):
                return f"页面提示保存失败/校验失败：{joined}"

        try:
            drawer = _get_person_detail_drawer(driver, "")
            if drawer:
                text = str(
                    driver.execute_script(
                        """
                        const normalize = (value) => String(value || '')
                          .replace(/\\s+/g, ' ')
                          .trim();
                        return normalize(arguments[0].innerText || arguments[0].textContent);
                        """,
                        drawer,
                    )
                )
                if employee_no in text:
                    return f"详情页已显示新的工号 {employee_no}"
        except Exception:
            pass

        time.sleep(0.4)

    if last_messages:
        return f"未等到明确保存成功提示；最后页面提示：{'；'.join(last_messages)}"
    return "未等到明确保存成功提示。"


def _wait_for_hire_date_save_result(
    driver: webdriver.Chrome,
    hire_date: str,
    timeout: int = 12,
) -> str:
    """Wait for save feedback and verify the detail drawer displays 入职时间."""
    from create_new_person import _collect_visible_messages

    expected_display = _format_hire_date_for_input(hire_date)
    success_markers = ("保存成功", "修改成功", "更新成功", "操作成功", "success")
    failure_markers = (
        "校验失败",
        "失败",
        "错误",
        "不能为空",
        "请选择",
        "请输入",
        "已存在",
        "重复",
    )
    deadline = time.monotonic() + timeout
    last_messages: list[str] = []
    saw_success_message = ""

    while time.monotonic() < deadline:
        messages = _collect_visible_messages(driver, fast=True)
        if messages:
            last_messages = messages
            joined = "；".join(messages)
            if any(marker in joined for marker in failure_markers):
                return f"页面提示保存失败/校验失败：{joined}"
            if any(marker in joined for marker in success_markers):
                saw_success_message = joined

        try:
            drawer = _get_person_detail_drawer(driver, "")
            if drawer:
                text = str(
                    driver.execute_script(
                        """
                        const normalize = (value) => String(value || '')
                          .replace(/\\s+/g, ' ')
                          .trim();
                        return normalize(arguments[0].innerText || arguments[0].textContent);
                        """,
                        drawer,
                    )
                )
                if expected_display in text:
                    if saw_success_message:
                        return (
                            f"页面提示保存成功：{saw_success_message}，"
                            f"详情页已显示入职时间 {expected_display}"
                        )
                    return f"详情页已显示入职时间 {expected_display}"
        except Exception:
            pass

        time.sleep(0.4)

    if saw_success_message:
        return (
            f"页面提示保存成功：{saw_success_message}，"
            f"但详情页未确认入职时间为 {expected_display}"
        )
    if last_messages:
        return f"未等到明确保存成功提示；最后页面提示：{'；'.join(last_messages)}"
    return f"未等到明确保存成功提示，也未确认入职时间为 {expected_display}。"


def edit_person(
    driver: webdriver.Chrome,
    keyword: str = DEFAULT_EDIT_PERSON_KEYWORD,
    employee_no: str = "",
) -> str:
    """
    Search people by keyword, randomly edit one matching person's 工号.

    Args:
        driver: Existing Selenium Chrome driver already logged in to eTeams.
        keyword: Search keyword; defaults to ``xuyingtest``.
        employee_no: New employee number. If omitted, a unique value is generated.

    Returns:
        Chinese status text containing the selected row and edit outcome.
    """
    keyword = _normalize_text(keyword) or DEFAULT_EDIT_PERSON_KEYWORD
    employee_no = _normalize_text(employee_no) or _generate_employee_no()
    step_results: list[str] = []

    hr_result = _ensure_human_resources_page(driver)
    step_results.append(f"步骤1：{hr_result}")
    if not hr_result.startswith(("已在", "已进入")):
        return "未能编辑人员：" + " ".join(step_results)

    search_result = _search_people_by_keyword(driver, keyword)
    step_results.append(f"步骤2：{search_result}")
    if not search_result.startswith("已"):
        return "未能编辑人员：" + " ".join(step_results)

    rows = _wait_for_matching_person_rows(driver, keyword)
    if not rows:
        return (
            f"未能编辑人员：搜索 {keyword} 后未找到可点击的匹配人员记录。"
            + " ".join(step_results)
        )

    selected_row = random.choice(rows)
    selected_snippet = _describe_row(driver, selected_row)
    open_attempts, drawer = _open_person_detail_from_row(driver, selected_row, keyword)
    step_results.append(
        f"步骤3：从 {len(rows)} 条匹配结果中随机选择一条"
        f"{'（' + selected_snippet + '）' if selected_snippet else ''}，"
        f"打开详情尝试：{open_attempts or '已点击'}。"
    )

    if not drawer:
        return "未能编辑人员：点击人员记录后未检测到人员详情页弹出。" + " ".join(step_results)

    edit_button = ""
    deadline = time.monotonic() + 8
    while time.monotonic() < deadline and not edit_button:
        drawer = _get_person_detail_drawer(driver, keyword) or drawer
        edit_button = _click_edit_button(driver, drawer)
        if not edit_button:
            time.sleep(0.3)

    if not edit_button:
        return "未能编辑人员：人员详情页中未找到「编辑」按钮。" + " ".join(step_results)

    try:
        drawer = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: (
                fresh_drawer
                if (fresh_drawer := _get_person_detail_drawer(drv, keyword))
                and _edit_mode_is_active(drv, fresh_drawer)
                else False
            )
        )
    except TimeoutException:
        return (
            "未能编辑人员：点击编辑后未检测到「保存/返回」编辑状态。"
            + " ".join(step_results)
        )

    step_results.append(f"步骤4：已点击「{edit_button}」，页面已切换为「保存/返回」编辑状态。")

    # After clicking「编辑」the page is already in edit mode and shows
    #「保存 / 返回」. Do not click the top「基本信息」tab again here: on this page
    # that can reload the read-only detail view and make the 工号 input
    # disappear before we type into it.
    try:
        employee_input = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: _get_employee_no_input(drv, _get_person_detail_drawer(drv, keyword) or drawer)
        )
    except TimeoutException:
        return "未能编辑人员：编辑页「基本资料」中未找到可编辑的「工号」输入框。" + " ".join(
            step_results
        )

    final_value = _set_input_value(driver, employee_input, employee_no)
    if employee_no not in final_value:
        return (
            f"未能编辑人员：未确认工号输入框已修改为 {employee_no}"
            f"（当前值：{final_value or '空'}）。"
            + " ".join(step_results)
        )
    step_results.append(f"步骤5：已将「基本资料」中的工号修改为 {employee_no}。")

    drawer = _get_person_detail_drawer(driver, keyword) or drawer
    save_button = _click_save_button(driver, drawer)
    if not save_button:
        return "未能编辑人员：未找到编辑页中的「保存」按钮。" + " ".join(step_results)

    save_result = _wait_for_save_result(driver, employee_no)
    step_results.append(f"步骤6：已点击「{save_button}」，{save_result}。")
    if "失败" in save_result or "校验失败" in save_result:
        return "未能确认人员工号修改成功：" + " ".join(step_results)

    return (
        f"已完成随机人员编辑：搜索关键词 {keyword}，"
        f"选中记录{f'（{selected_snippet}）' if selected_snippet else ''}，"
        f"并将工号修改为 {employee_no}。"
        + " ".join(step_results)
    )


def edit_person_hire_date(
    driver: webdriver.Chrome,
    keyword: str = DEFAULT_EDIT_PERSON_KEYWORD,
    hire_date: str = "",
) -> str:
    """
    Search people by keyword, randomly edit one matching person's 入职时间.

    The date input is clicked, then a yyyyMMdd value (for example 20020211)
    is typed directly and saved through the page's normal「保存」button.

    Args:
        driver: Existing Selenium Chrome driver already logged in to eTeams.
        keyword: Search keyword; defaults to ``xuyingtest``.
        hire_date: Date text in yyyyMMdd format. If omitted, a random valid
            date is generated.

    Returns:
        Chinese status text containing the selected row and edit outcome.
    """
    keyword = _normalize_text(keyword) or DEFAULT_EDIT_PERSON_KEYWORD
    hire_date = _normalize_text(hire_date) or _generate_hire_date_value()
    step_results: list[str] = []

    hr_result = _ensure_human_resources_page(driver)
    step_results.append(f"步骤1：{hr_result}")
    if not hr_result.startswith(("已在", "已进入")):
        return "未能编辑人员入职时间：" + " ".join(step_results)

    search_result = _search_people_by_keyword(driver, keyword)
    step_results.append(f"步骤2：{search_result}")
    if not search_result.startswith("已"):
        return "未能编辑人员入职时间：" + " ".join(step_results)

    rows = _wait_for_matching_person_rows(driver, keyword)
    if not rows:
        return (
            f"未能编辑人员入职时间：搜索 {keyword} 后未找到可点击的匹配人员记录。"
            + " ".join(step_results)
        )

    selected_row = random.choice(rows)
    selected_snippet = _describe_row(driver, selected_row)
    open_attempts, drawer = _open_person_detail_from_row(driver, selected_row, keyword)
    step_results.append(
        f"步骤3：从 {len(rows)} 条匹配结果中随机选择一条"
        f"{'（' + selected_snippet + '）' if selected_snippet else ''}，"
        f"打开详情尝试：{open_attempts or '已点击'}。"
    )

    if not drawer:
        return "未能编辑人员入职时间：点击人员记录后未检测到人员详情页弹出。" + " ".join(
            step_results
        )

    edit_button = ""
    deadline = time.monotonic() + 8
    while time.monotonic() < deadline and not edit_button:
        drawer = _get_person_detail_drawer(driver, keyword) or drawer
        edit_button = _click_edit_button(driver, drawer)
        if not edit_button:
            time.sleep(0.3)

    if not edit_button:
        return "未能编辑人员入职时间：人员详情页中未找到「编辑」按钮。" + " ".join(step_results)

    try:
        drawer = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: (
                fresh_drawer
                if (fresh_drawer := _get_person_detail_drawer(drv, keyword))
                and _edit_mode_is_active(drv, fresh_drawer)
                else False
            )
        )
    except TimeoutException:
        return (
            "未能编辑人员入职时间：点击编辑后未检测到「保存/返回」编辑状态。"
            + " ".join(step_results)
        )

    step_results.append(f"步骤4：已点击「{edit_button}」，页面已切换为「保存/返回」编辑状态。")

    try:
        date_input = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: _get_hire_date_input(
                drv,
                _get_person_detail_drawer(drv, keyword) or drawer,
            )
        )
    except TimeoutException:
        return (
            "未能编辑人员入职时间：编辑页「基本资料」中未找到 placeholder 为「选择日期」"
            "的「入职时间」输入框。"
            + " ".join(step_results)
        )

    final_date_value = _type_hire_date_value(driver, date_input, hire_date)
    if not final_date_value or not _hire_date_values_equal(final_date_value, hire_date):
        return (
            f"未能编辑人员入职时间：已点击「选择日期」输入框，但未能输入日期 {hire_date}。"
            f"（当前值：{final_date_value or '空'}）"
            + " ".join(step_results)
        )

    step_results.append(
        f"步骤5：已点击「入职时间」日期输入框，直接输入 {hire_date}，"
        f"输入框当前值为 {final_date_value}。"
    )

    drawer = _get_person_detail_drawer(driver, keyword) or drawer
    save_button = _click_save_button(driver, drawer)
    if not save_button:
        return "未能编辑人员入职时间：未找到编辑页中的「保存」按钮。" + " ".join(step_results)

    save_result = _wait_for_hire_date_save_result(driver, final_date_value)
    step_results.append(f"步骤6：已点击「{save_button}」，{save_result}。")
    if "失败" in save_result or "校验失败" in save_result or "未确认" in save_result:
        return "未能确认人员入职时间修改成功：" + " ".join(step_results)

    return (
        f"已完成随机人员入职时间编辑：搜索关键词 {keyword}，"
        f"选中记录{f'（{selected_snippet}）' if selected_snippet else ''}，"
        f"并将入职时间修改为 {final_date_value}。"
        + " ".join(step_results)
    )


def edit_person_employee_no_and_hire_date(
    driver: webdriver.Chrome,
    keyword: str = DEFAULT_EDIT_PERSON_KEYWORD,
    employee_no: str = "",
    hire_date: str = "",
) -> str:
    """
    Search people by keyword, randomly edit one person's 工号 and 入职时间.

    Both fields are changed in the same edit session, then saved once.

    Args:
        driver: Existing Selenium Chrome driver already logged in to eTeams.
        keyword: Search keyword; defaults to ``xuyingtest``.
        employee_no: New employee number. If omitted, a unique value is generated.
        hire_date: Date text in yyyyMMdd format. If omitted, a random valid
            date is generated.

    Returns:
        Chinese status text containing the selected row and edit outcome.
    """
    keyword = _normalize_text(keyword) or DEFAULT_EDIT_PERSON_KEYWORD
    employee_no = _normalize_text(employee_no) or _generate_employee_no()
    hire_date = _normalize_text(hire_date) or _generate_hire_date_value()
    step_results: list[str] = []

    hr_result = _ensure_human_resources_page(driver)
    step_results.append(f"步骤1：{hr_result}")
    if not hr_result.startswith(("已在", "已进入")):
        return "未能同时编辑人员工号和入职时间：" + " ".join(step_results)

    search_result = _search_people_by_keyword(driver, keyword)
    step_results.append(f"步骤2：{search_result}")
    if not search_result.startswith("已"):
        return "未能同时编辑人员工号和入职时间：" + " ".join(step_results)

    rows = _wait_for_matching_person_rows(driver, keyword)
    if not rows:
        return (
            f"未能同时编辑人员工号和入职时间：搜索 {keyword} 后未找到可点击的匹配人员记录。"
            + " ".join(step_results)
        )

    selected_row = random.choice(rows)
    selected_snippet = _describe_row(driver, selected_row)
    open_attempts, drawer = _open_person_detail_from_row(driver, selected_row, keyword)
    step_results.append(
        f"步骤3：从 {len(rows)} 条匹配结果中随机选择一条"
        f"{'（' + selected_snippet + '）' if selected_snippet else ''}，"
        f"打开详情尝试：{open_attempts or '已点击'}。"
    )

    if not drawer:
        return (
            "未能同时编辑人员工号和入职时间：点击人员记录后未检测到人员详情页弹出。"
            + " ".join(step_results)
        )

    edit_button = ""
    deadline = time.monotonic() + 8
    while time.monotonic() < deadline and not edit_button:
        drawer = _get_person_detail_drawer(driver, keyword) or drawer
        edit_button = _click_edit_button(driver, drawer)
        if not edit_button:
            time.sleep(0.3)

    if not edit_button:
        return (
            "未能同时编辑人员工号和入职时间：人员详情页中未找到「编辑」按钮。"
            + " ".join(step_results)
        )

    try:
        drawer = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: (
                fresh_drawer
                if (fresh_drawer := _get_person_detail_drawer(drv, keyword))
                and _edit_mode_is_active(drv, fresh_drawer)
                else False
            )
        )
    except TimeoutException:
        return (
            "未能同时编辑人员工号和入职时间：点击编辑后未检测到「保存/返回」编辑状态。"
            + " ".join(step_results)
        )

    step_results.append(f"步骤4：已点击「{edit_button}」，页面已切换为「保存/返回」编辑状态。")

    try:
        employee_input = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: _get_employee_no_input(
                drv,
                _get_person_detail_drawer(drv, keyword) or drawer,
            )
        )
    except TimeoutException:
        return (
            "未能同时编辑人员工号和入职时间：编辑页「基本资料」中未找到可编辑的「工号」输入框。"
            + " ".join(step_results)
        )

    final_employee_no = _set_input_value(driver, employee_input, employee_no)
    if employee_no not in final_employee_no:
        return (
            f"未能同时编辑人员工号和入职时间：未确认工号输入框已修改为 {employee_no}"
            f"（当前值：{final_employee_no or '空'}）。"
            + " ".join(step_results)
        )
    step_results.append(f"步骤5：已将「基本资料」中的工号修改为 {employee_no}。")

    try:
        date_input = WebDriverWait(driver, 8, poll_frequency=0.3).until(
            lambda drv: _get_hire_date_input(
                drv,
                _get_person_detail_drawer(drv, keyword) or drawer,
            )
        )
    except TimeoutException:
        return (
            "未能同时编辑人员工号和入职时间：编辑页「基本资料」中未找到 placeholder 为「选择日期」"
            "的「入职时间」输入框。"
            + " ".join(step_results)
        )

    final_date_value = _type_hire_date_value(driver, date_input, hire_date)
    if not final_date_value or not _hire_date_values_equal(final_date_value, hire_date):
        return (
            f"未能同时编辑人员工号和入职时间：已点击「选择日期」输入框，"
            f"但未能输入日期 {hire_date}（当前值：{final_date_value or '空'}）。"
            + " ".join(step_results)
        )
    step_results.append(
        f"步骤6：已点击「入职时间」日期输入框，直接输入 {hire_date}，"
        f"输入框当前值为 {final_date_value}。"
    )

    drawer = _get_person_detail_drawer(driver, keyword) or drawer
    save_button = _click_save_button(driver, drawer)
    if not save_button:
        return "未能同时编辑人员工号和入职时间：未找到编辑页中的「保存」按钮。" + " ".join(
            step_results
        )

    save_result = _wait_for_hire_date_save_result(driver, final_date_value)
    step_results.append(f"步骤7：已点击「{save_button}」，{save_result}。")
    if "失败" in save_result or "校验失败" in save_result or "未确认" in save_result:
        return "未能确认人员工号和入职时间同时修改成功：" + " ".join(step_results)

    return (
        f"已完成随机人员工号和入职时间同时编辑：搜索关键词 {keyword}，"
        f"选中记录{f'（{selected_snippet}）' if selected_snippet else ''}，"
        f"工号修改为 {employee_no}，入职时间修改为 {final_date_value}。"
        + " ".join(step_results)
    )


__all__ = [
    "edit_person",
    "edit_person_hire_date",
    "edit_person_employee_no_and_hire_date",
]
