---
name: field-plus-selector
description: Reusable Selenium automation pattern for web form fields where a visible label is followed by a circular plus/search selector, such as eTeams fields like 部门、岗位、职级、职称、所属机构、办公地点. Use when Codex needs to add, debug, or refactor automation that clicks the + control, selects a dropdown/popup candidate, verifies the value is written back, or handles switch-controlled optional selector fields.
---

# Field Plus Selector

## Overview

Use this skill for the common UI pattern: a field label in a form row followed by a `+` / search / browser selector control. The bundled helper avoids brittle absolute coordinates by finding the label, scanning the same row for the selector control, choosing a random candidate, and verifying the selected value is displayed back in that row.

Bundled helper: `scripts/field_plus_selector.py`.

## Workflow

1. Check whether the target project already has `field_plus_selector.py`.
   - If absent, copy this skill's `scripts/field_plus_selector.py` into the project.
   - If present, patch the project copy instead of duplicating another helper.
2. Import the helper where the Selenium flow fills the form:

```python
from field_plus_selector import fill_field_plus_selector
```

3. Provide a `root_getter` when the form is inside a drawer/dialog. The root must return the visible dialog element so the helper ignores the page menu/tree/table behind it.

```python
def _get_visible_dialog(driver):
    return driver.execute_script("""
      const visible = (el) => {
        if (!el) return false;
        const style = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        return style.display !== 'none'
          && style.visibility !== 'hidden'
          && rect.width > 0
          && rect.height > 0;
      };
      return Array.from(document.querySelectorAll('.ui-dialog-wrap-right, .ui-dialog-wrap'))
        .filter((el) => visible(el) && String(el.innerText || '').includes('新建'))
        .sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left)[0] || null;
    """)
```

4. Fill required selector fields with `optional=False`:

```python
result = fill_field_plus_selector(
    driver,
    "部门",
    root_getter=_get_visible_dialog,
)
if not result.startswith("部门="):
    return f"部门信息填写失败：{result}"
```

5. Fill switch-controlled fields with `optional=True`. If the switch is off and the field is absent, treat the skip result as successful and continue saving.

```python
organization_result = fill_field_plus_selector(
    driver,
    "所属机构",
    root_getter=_get_visible_dialog,
    optional=True,
)
if not organization_result.startswith(("所属机构=", "所属机构未启用")):
    return f"所属机构填写失败：{organization_result}"
```

6. Validate with `python -m py_compile` and at least one real Selenium run for each enabled/disabled optional-field state when possible.

## Helper Behavior

- Locates the label by exact or contained visible text, favoring the smallest exact label node.
- Clicks compact selector controls in the same row to the right of the label, including `.associative-search-icon`, `.ui-input-suffix`, `Icon-add-to01`, semantic buttons, and plain compact plus elements.
- Chooses a random visible candidate from common dropdown/list/tree option containers.
- Avoids selecting disabled, empty, placeholder, and `无数据` options.
- Avoids repeating the same candidate for the same field when alternatives exist.
- Verifies the chosen value appears back in the field row.

## Guardrails

- Do not click the first visible `+` on the page globally; always scope by field label and preferably by dialog/drawer root.
- Do not hardcode screen coordinates except as a last-resort debug step.
- Do not fail optional switch-controlled fields when absent; use `optional=True`.
- Do not delete or modify existing business data while validating selector filling.
