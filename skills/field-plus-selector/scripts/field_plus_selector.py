"""
Reusable Selenium helper for the common eTeams "field label + plus button"
selector pattern.

Typical UI:

    部门        (+)
    岗位        (+)
    所属机构    (+)

The helper finds a visible field label, clicks the selector/+ control in the
same row, randomly chooses a visible candidate from the opened popup/dropdown,
and verifies the selected value is written back to the row.
"""

from __future__ import annotations

import secrets
import time
from collections.abc import Callable
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

RootGetter = Callable[[webdriver.Chrome], Any]

_LAST_SELECTED_BY_FIELD: dict[str, str] = {}


def field_plus_selected_text(
    driver: webdriver.Chrome,
    field_label: str,
    *,
    root_getter: RootGetter | None = None,
) -> str:
    """Return selected text in a field row, or an empty string if none is shown."""
    try:
        return str(
            driver.execute_script(
                """
                const root = arguments[0] || document.body;
                const fieldLabel = arguments[1];
                if (!root) return '';
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                const cleanLabel = (value) => normalize(value)
                  .replace(/[\\*＊:：]/g, '')
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
                const findLabel = (root, fieldLabel) => {
                  const fieldClean = cleanLabel(fieldLabel);
                  return Array.from(root.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))
                    .map((label) => {
                      if (!visible(label)) return null;
                      const rawText = normalize(
                        label.innerText
                        || label.textContent
                        || label.getAttribute('title')
                        || label.getAttribute('aria-label')
                      );
                      const text = cleanLabel(rawText);
                      if (!text) return null;
                      const exact = text === fieldClean;
                      const contains = text.includes(fieldClean);
                      if (!exact && !contains) return null;
                      const rect = label.getBoundingClientRect();
                      const area = Math.max(1, rect.width * rect.height);
                      const score = (exact ? 100000 : 0)
                        + (text.startsWith(fieldClean) ? 1000 : 0)
                        - area / 1000
                        - Math.max(0, text.length - fieldClean.length) * 10;
                      return { element: label, rect, score, text };
                    })
                    .filter(Boolean)
                    .sort((a, b) => b.score - a.score || a.rect.top - b.rect.top || a.rect.left - b.rect.left)[0]?.element || null;
                };

                const label = findLabel(root, fieldLabel);
                if (!label) return '';
                const labelRect = label.getBoundingClientRect();
                const labelCenterY = (labelRect.top + labelRect.bottom) / 2;
                const fieldClean = cleanLabel(fieldLabel);
                const emptyText = /请选择|请输入|无数据|暂无数据|please select|please enter|select|choose|no data/i;
                const otherLabels = /群组名称|群组类型|显示顺序|群组编号|Group Name|Group Type|Display Order|Group No/i;
                const selected = [];
                for (const el of Array.from(root.querySelectorAll([
                  '.ui-browser-associative-selected-item',
                  '.ui-select-selection-item',
                  '.ui-select-input-selected',
                  '.ant-select-selection-item',
                  '.el-select__tags-text',
                  '.ui-tag',
                  '[title]',
                  'input',
                  'span',
                  'div'
                ].join(',')))) {
                  if (!visible(el)) continue;
                  if (el === label || el.contains(label)) continue;
                  const rect = el.getBoundingClientRect();
                  const centerY = (rect.top + rect.bottom) / 2;
                  if (centerY < labelCenterY - 32 || centerY > labelCenterY + 32) continue;
                  if (rect.left <= labelRect.right - 4) continue;
                  const text = normalize(
                    el.value
                    || el.getAttribute('title')
                    || el.getAttribute('aria-label')
                    || el.innerText
                    || el.textContent
                  );
                  if (!text || emptyText.test(text)) continue;
                  if (/^\\+$/.test(text)) continue;
                  let selectedText = cleanLabel(text);
                  if (!selectedText || selectedText === fieldClean) continue;
                  if (otherLabels.test(selectedText)) continue;
                  if (selectedText.includes(fieldLabel)) {
                    selectedText = cleanLabel(selectedText.replace(fieldLabel, ''));
                    if (!selectedText) continue;
                  }
                  if (selectedText.length > 100) continue;
                  selected.push(selectedText);
                }
                return selected[0] || '';
                """,
                root_getter(driver) if root_getter else None,
                field_label,
            )
        ).strip()
    except Exception:
        return ""


def field_plus_is_present(
    driver: webdriver.Chrome,
    field_label: str,
    *,
    root_getter: RootGetter | None = None,
) -> bool:
    """Return True when a field row exists and has a visible plus/selector control."""
    try:
        return bool(
            driver.execute_script(
                """
                const root = arguments[0] || document.body;
                const fieldLabel = arguments[1];
                if (!root) return false;
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                const cleanLabel = (value) => normalize(value)
                  .replace(/[\\*＊:：]/g, '')
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
                const findLabel = (root, fieldLabel) => {
                  const fieldClean = cleanLabel(fieldLabel);
                  return Array.from(root.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))
                    .map((label) => {
                      if (!visible(label)) return null;
                      const rawText = normalize(
                        label.innerText
                        || label.textContent
                        || label.getAttribute('title')
                        || label.getAttribute('aria-label')
                      );
                      const text = cleanLabel(rawText);
                      if (!text) return null;
                      const exact = text === fieldClean;
                      const contains = text.includes(fieldClean);
                      if (!exact && !contains) return null;
                      const rect = label.getBoundingClientRect();
                      const area = Math.max(1, rect.width * rect.height);
                      const score = (exact ? 100000 : 0)
                        + (text.startsWith(fieldClean) ? 1000 : 0)
                        - area / 1000
                        - Math.max(0, text.length - fieldClean.length) * 10;
                      return { element: label, rect, score, text };
                    })
                    .filter(Boolean)
                    .sort((a, b) => b.score - a.score || a.rect.top - b.rect.top || a.rect.left - b.rect.left)[0]?.element || null;
                };

                const label = findLabel(root, fieldLabel);
                if (!label) return false;
                const labelRect = label.getBoundingClientRect();
                const labelCenterY = (labelRect.top + labelRect.bottom) / 2;
                const controlSelector = [
                  '.associative-search-icon',
                  '.ui-input-suffix',
                  '.Icon-add-to01',
                  '[class*="add"]',
                  '[class*="Add"]',
                  'button',
                  '[role="button"]',
                  'input[placeholder="请选择"]',
                  '.ui-input-wrap',
                  '.ui-browser-associative-search',
                  '.ui-browser',
                  '.ui-select',
                  '.ant-select',
                  '.el-select',
                  'input'
                ].join(',');
                for (const control of Array.from(root.querySelectorAll(controlSelector))) {
                  if (!visible(control)) continue;
                  const rect = control.getBoundingClientRect();
                  const centerY = (rect.top + rect.bottom) / 2;
                  if (centerY < labelCenterY - 32 || centerY > labelCenterY + 32) continue;
                  if (rect.left <= labelRect.right - 4) continue;
                  if (rect.width > 700 || rect.height > 90) continue;
                  return true;
                }
                return false;
                """,
                root_getter(driver) if root_getter else None,
                field_label,
            )
        )
    except Exception:
        return False


def click_field_plus_control(
    driver: webdriver.Chrome,
    field_label: str,
    *,
    root_getter: RootGetter | None = None,
) -> str:
    """Click the plus/selector control in the same row as ``field_label``."""
    try:
        return str(
            driver.execute_script(
                """
                const root = arguments[0] || document.body;
                const fieldLabel = arguments[1];
                if (!root) return '';
                const normalize = (value) => String(value || '')
                  .replace(/\\s+/g, ' ')
                  .trim();
                const cleanLabel = (value) => normalize(value)
                  .replace(/[\\*＊:：]/g, '')
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
                const findLabel = (root, fieldLabel) => {
                  const fieldClean = cleanLabel(fieldLabel);
                  return Array.from(root.querySelectorAll('.ui-formItem-label-span, [title], label, span, div'))
                    .map((label) => {
                      if (!visible(label)) return null;
                      const rawText = normalize(
                        label.innerText
                        || label.textContent
                        || label.getAttribute('title')
                        || label.getAttribute('aria-label')
                      );
                      const text = cleanLabel(rawText);
                      if (!text) return null;
                      const exact = text === fieldClean;
                      const contains = text.includes(fieldClean);
                      if (!exact && !contains) return null;
                      const rect = label.getBoundingClientRect();
                      const area = Math.max(1, rect.width * rect.height);
                      const score = (exact ? 100000 : 0)
                        + (text.startsWith(fieldClean) ? 1000 : 0)
                        - area / 1000
                        - Math.max(0, text.length - fieldClean.length) * 10;
                      return { element: label, rect, score, text };
                    })
                    .filter(Boolean)
                    .sort((a, b) => b.score - a.score || a.rect.top - b.rect.top || a.rect.left - b.rect.left)[0]?.element || null;
                };

                const label = findLabel(root, fieldLabel);
                if (!label) return '';
                const labelRect = label.getBoundingClientRect();
                const labelCenterY = (labelRect.top + labelRect.bottom) / 2;
                const candidates = [];
                const pushCandidate = (raw, baseScore = 0) => {
                  if (!visible(raw)) return;
                  const clickable = raw.closest([
                    'button',
                    '[role="button"]',
                    '.ui-input-suffix',
                    '.associative-search-icon',
                    '.ui-input-wrap',
                    '.ui-browser-associative-search',
                    '.ui-browser',
                    '.ui-select',
                    '.ui-select-selector',
                    '.ant-select',
                    '.ant-select-selector',
                    '.el-select'
                  ].join(',')) || raw;
                  if (!visible(clickable)) return;
                  const rect = clickable.getBoundingClientRect();
                  const centerY = (rect.top + rect.bottom) / 2;
                  if (centerY < labelCenterY - 32 || centerY > labelCenterY + 32) return;
                  if (rect.left <= labelRect.right - 4) return;
                  if (rect.width > 700 || rect.height > 90) return;
                  const className = String(clickable.className || '');
                  const text = normalize(
                    raw.innerText
                    || raw.textContent
                    || raw.getAttribute('title')
                    || raw.getAttribute('aria-label')
                    || ''
                  );
                  const attrs = `${className} ${raw.tagName || ''} ${raw.getAttribute('placeholder') || ''} ${text}`.toLowerCase();
                  let score = baseScore;
                  if (/^\\+$/.test(text)) score += 150;
                  if (/associative-search-icon|ui-input-suffix|icon-add-to01|add|plus/.test(attrs)) score += 100;
                  if (/select|browser|请选择/.test(attrs)) score += 40;
                  if (raw.matches && raw.matches('input[placeholder="请选择"]')) score += 30;
                  score += Math.max(0, 80 - Math.max(rect.width, rect.height)) / 10;
                  score += rect.left / 1000;
                  score -= Math.abs(centerY - labelCenterY) / 100;
                  candidates.push({ element: clickable, score, className, left: rect.left, top: rect.top });
                };

                const selector = [
                  '.associative-search-icon',
                  '.ui-input-suffix',
                  '.Icon-add-to01',
                  '[class*="add"]',
                  '[class*="Add"]',
                  'button',
                  '[role="button"]',
                  'svg',
                  'use',
                  'input[placeholder="请选择"]',
                  '.ui-input-wrap',
                  '.ui-browser-associative-search',
                  '.ui-browser',
                  '.ui-select',
                  '.ui-select-selector',
                  '.ant-select',
                  '.ant-select-selector',
                  '.el-select',
                  'input'
                ].join(',');
                for (const raw of Array.from(root.querySelectorAll(selector))) {
                  pushCandidate(raw, 0);
                }
                for (const raw of Array.from(root.querySelectorAll('*'))) {
                  const rect = raw.getBoundingClientRect();
                  if (rect.width < 8 || rect.height < 8) continue;
                  if (rect.width > 90 || rect.height > 70) continue;
                  pushCandidate(raw, 10);
                }

                candidates.sort((a, b) => b.score - a.score || b.left - a.left || a.top - b.top);
                let target = candidates[0]?.element;
                if (!target) {
                  const rootRect = root.getBoundingClientRect ? root.getBoundingClientRect() : document.body.getBoundingClientRect();
                  for (let x = labelRect.right + 20; x < Math.min(rootRect.right - 20, labelRect.right + 700); x += 20) {
                    const raw = document.elementFromPoint(x, labelCenterY);
                    if (!raw || !root.contains(raw) || !visible(raw)) continue;
                    const clickable = raw.closest('button, [role="button"], .ui-input-suffix, .associative-search-icon, .ui-input-wrap, .ui-browser-associative-search, .ui-browser, .ui-select, .ui-select-selector, .ant-select, .ant-select-selector, .el-select') || raw;
                    if (!visible(clickable)) continue;
                    const rect = clickable.getBoundingClientRect();
                    if (rect.left <= labelRect.right - 4 || rect.width > 700 || rect.height > 90) continue;
                    target = clickable;
                    break;
                  }
                }
                if (!target) return '';
                const clickedLabel = candidates[0]?.className || `${fieldLabel}选择控件`;
                fireClick(target);
                return normalize(clickedLabel) || `${fieldLabel}选择控件`;
                """,
                root_getter(driver) if root_getter else None,
                field_label,
            )
        ).strip()
    except Exception:
        return ""


def click_random_plus_dropdown_candidate(
    driver: webdriver.Chrome,
    field_label: str,
    *,
    root_getter: RootGetter | None = None,
) -> str:
    """Click a random visible candidate opened by a field-plus selector."""
    root = root_getter(driver) if root_getter else None
    candidates = driver.execute_script(
        """
        const root = arguments[0];
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
        const rootRect = root?.getBoundingClientRect?.();
        const badText = /无数据|暂无数据|没有数据|请选择|请输入|No Data|No options|Select|Search/i;
        const selector = [
          '.ui-browser-associative-dropdown-list .ui-list-item',
          '.ui-browser-associative-dropdown-list-item',
          '.ui-select-dropdown .ui-list-item',
          '.ui-dropdown .ui-list-item',
          '.ui-list-scrollview .ui-list-item',
          '.ant-select-dropdown [role="option"]',
          '.el-select-dropdown .el-select-dropdown__item',
          '.rc-virtual-list-holder-inner [role="option"]',
          '[role="option"]',
          '[role="treeitem"]'
        ].join(',');
        const candidates = [];
        const seen = new Set();
        for (const raw of Array.from(document.body.querySelectorAll(selector))) {
          if (!visible(raw)) continue;
          if (root && (raw === root || root.contains(raw))) continue;
          const text = normalize(
            raw.innerText
            || raw.textContent
            || raw.getAttribute('title')
            || raw.getAttribute('aria-label')
          );
          if (!text || badText.test(text) || text.length > 100) continue;
          const rect = raw.getBoundingClientRect();
          if (rootRect) {
            if (rect.top < Math.max(80, rootRect.top - 20)) continue;
            if (rect.right < rootRect.left - 120 && rect.left < 500) continue;
          } else if (rect.left < 500 || rect.top < 80) {
            continue;
          }
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
        root,
    )
    candidates = [str(candidate).strip() for candidate in (candidates or []) if str(candidate).strip()]
    if not candidates:
        return ""

    last_selected = _LAST_SELECTED_BY_FIELD.get(field_label, "")
    eligible = [candidate for candidate in candidates if candidate != last_selected] or candidates
    selected_text = eligible[secrets.randbelow(len(eligible))]
    selected_index = candidates.index(selected_text)

    clicked = str(
        driver.execute_script(
            """
            const root = arguments[0];
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
            const rootRect = root?.getBoundingClientRect?.();
            const badText = /无数据|暂无数据|没有数据|请选择|请输入|No Data|No options|Select|Search/i;
            const selector = [
              '.ui-browser-associative-dropdown-list .ui-list-item',
              '.ui-browser-associative-dropdown-list-item',
              '.ui-select-dropdown .ui-list-item',
              '.ui-dropdown .ui-list-item',
              '.ui-list-scrollview .ui-list-item',
              '.ant-select-dropdown [role="option"]',
              '.el-select-dropdown .el-select-dropdown__item',
              '.rc-virtual-list-holder-inner [role="option"]',
              '[role="option"]',
              '[role="treeitem"]'
            ].join(',');
            const candidates = [];
            const seen = new Set();
            for (const raw of Array.from(document.body.querySelectorAll(selector))) {
              if (!visible(raw)) continue;
              if (root && (raw === root || root.contains(raw))) continue;
              const text = normalize(
                raw.innerText
                || raw.textContent
                || raw.getAttribute('title')
                || raw.getAttribute('aria-label')
              );
              if (!text || badText.test(text) || text.length > 100) continue;
              const rect = raw.getBoundingClientRect();
              if (rootRect) {
                if (rect.top < Math.max(80, rootRect.top - 20)) continue;
                if (rect.right < rootRect.left - 120 && rect.left < 500) continue;
              } else if (rect.left < 500 || rect.top < 80) {
                continue;
              }
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
            root,
            selected_index,
        )
    ).strip()
    if clicked:
        _LAST_SELECTED_BY_FIELD[field_label] = selected_text
    return clicked


def fill_field_plus_selector(
    driver: webdriver.Chrome,
    field_label: str,
    *,
    root_getter: RootGetter | None = None,
    optional: bool = False,
    presence_timeout: float = 5.0,
    selection_timeout: float = 8.0,
    max_attempts: int = 3,
) -> str:
    """
    Fill a "field label + plus" selector by randomly choosing one candidate.

    Returns:
        - ``"{field_label}=value"`` on success.
        - ``"{field_label}未启用，已跳过。"`` when optional=True and absent.
        - A Chinese failure reason otherwise.
    """
    try:
        WebDriverWait(driver, presence_timeout, poll_frequency=0.3).until(
            lambda drv: field_plus_is_present(drv, field_label, root_getter=root_getter)
        )
    except TimeoutException:
        if optional:
            return f"{field_label}未启用，已跳过。"
        return f"未找到「{field_label}」字段或其选择控件。"

    existing = field_plus_selected_text(driver, field_label, root_getter=root_getter)
    if existing:
        return f"{field_label}={existing}"

    clicked = click_field_plus_control(driver, field_label, root_getter=root_getter)
    if not clicked:
        return f"未能点击「{field_label}」字段后的选择控件。"

    selected = ""
    for _attempt in range(max(1, int(max_attempts))):
        try:
            selected = WebDriverWait(driver, selection_timeout, poll_frequency=0.3).until(
                lambda drv: click_random_plus_dropdown_candidate(
                    drv,
                    field_label,
                    root_getter=root_getter,
                )
            )
            break
        except TimeoutException:
            clicked = click_field_plus_control(driver, field_label, root_getter=root_getter)
            if not clicked:
                break
            time.sleep(0.3)

    if not selected:
        return f"已点击「{field_label}」选择控件，但未能选择候选项。"

    try:
        confirmed_value = WebDriverWait(driver, selection_timeout, poll_frequency=0.3).until(
            lambda drv: field_plus_selected_text(drv, field_label, root_getter=root_getter)
        )
    except TimeoutException:
        confirmed_value = ""

    if not confirmed_value:
        return f"已随机选择 {selected}，但未能在「{field_label}」字段中确认已回填。"

    if "随机候选" in selected and confirmed_value in selected:
        return f"{field_label}={selected}"
    return f"{field_label}={confirmed_value}"


__all__ = [
    "click_field_plus_control",
    "click_random_plus_dropdown_candidate",
    "field_plus_is_present",
    "field_plus_selected_text",
    "fill_field_plus_selector",
]
