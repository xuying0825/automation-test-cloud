"""
Standalone Selenium helper for creating a new department in eTeams.

Usage:
    from create_new_department import create_new_department
    result = create_new_department(driver, department_name="xuyingtest001")

Precondition: ``driver`` is already logged in to eTeams and on the Org. Structure page.
"""

import time
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

DEFAULT_DEPARTMENT_PREFIX = "xuyingtest"


def _wait_for_element(driver, by, value, timeout=10):
    """Wait for an element to be present and visible."""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, value))
        )
    except TimeoutException:
        return None


def _click_element(driver, element):
    """Click an element using JavaScript to avoid interception issues."""
    try:
        driver.execute_script("arguments[0].click();", element)
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"点击元素失败: {e}")
        return False


def _switch_to_division_info_tab(driver):
    """Switch to the Division Info (分部信息) tab."""
    try:
        # Try to find the tab by text content
        tabs = driver.find_elements(By.CSS_SELECTOR, ".ant-tabs-tab, .tabs-tab, [role='tab']")
        
        for tab in tabs:
            tab_text = tab.text.strip()
            if "分部信息" in tab_text or "Division Info" in tab_text:
                # Check if already active
                if "active" not in tab.get_attribute("class", "").lower():
                    _click_element(driver, tab)
                    time.sleep(1)
                return True, f"已切换到「分部信息」页签"
        
        return False, "未找到「分部信息」页签"
    except Exception as e:
        return False, f"切换页签失败: {str(e)}"


def _click_new_department_button(driver):
    """Click the New Department button."""
    try:
        # Try multiple selectors for the new department button
        selectors = [
            "button[contains(text(), '新建部门')]",
            "button:contains('新建部门')",
            "[data-action='add']",
            ".ant-btn-primary",
            "button.ant-btn-primary",
        ]
        
        # First try using XPath with text
        buttons = driver.find_elements(By.XPATH, "//button[contains(text(), '新建部门')]")
        if buttons:
            _click_element(driver, buttons[0])
            time.sleep(1)
            return True, "已点击「新建部门」按钮"
        
        # Try other selectors
        for selector in selectors:
            try:
                btn = driver.find_element(By.CSS_SELECTOR, selector)
                if btn.is_displayed():
                    _click_element(driver, btn)
                    time.sleep(1)
                    return True, "已点击「新建部门」按钮"
            except NoSuchElementException:
                continue
        
        return False, "未找到「新建部门」按钮"
    except Exception as e:
        return False, f"点击新建部门按钮失败: {str(e)}"


def _fill_department_form(driver, department_name):
    """Fill in the department name only."""
    try:
        # Wait for the dialog to appear
        time.sleep(2)
        
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        
        # Find department name input field
        name_input = None
        
        # Strategy 1: Find input with class="ui-input" and placeholder="请输入"
        try:
            inputs = driver.find_elements(By.CSS_SELECTOR, "input.ui-input[placeholder='请输入']")
            for inp in inputs:
                if inp.is_displayed() and not inp.get_attribute('readonly') and not inp.get_attribute('disabled'):
                    name_input = inp
                    break
        except Exception:
            pass
        
        # Strategy 2: Find any visible ui-input
        if not name_input:
            try:
                inputs = driver.find_elements(By.CSS_SELECTOR, "input.ui-input")
                for inp in inputs:
                    if (inp.is_displayed() and 
                        not inp.get_attribute('readonly') and 
                        not inp.get_attribute('disabled')):
                        name_input = inp
                        break
            except Exception:
                pass
        
        if not name_input:
            return False, "未找到部门名称输入框", None, None, None
        
        # Fill in department name using reliable method
        ActionChains(driver).click(name_input).perform()
        time.sleep(0.3)
        
        # Clear the field
        try:
            name_input.clear()
        except Exception:
            name_input.send_keys(Keys.CONTROL, "a")
            name_input.send_keys(Keys.BACKSPACE)
        
        # Type the department name
        name_input.send_keys(department_name)
        time.sleep(0.3)
        
        # Trigger events using JavaScript to ensure the value is registered
        driver.execute_script(
            """
            const input = arguments[0];
            const value = arguments[1];
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
            if (setter) setter.call(input, value); else input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            """,
            name_input,
            department_name,
        )
        
        return True, f"已输入部门名称: {department_name}", None, None, None
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, f"填写表单失败: {str(e)}", None, None, None


def _click_save_button(driver):
    """Click the Save button in the dialog."""
    try:
        # Wait for save button to appear
        time.sleep(0.5)
        
        # Try to find save button using XPath with text
        buttons = driver.find_elements(By.XPATH, "//button[contains(text(), '保存') or contains(text(), '确定') or contains(text(), 'Save') or contains(text(), 'OK')]")
        if buttons:
            for btn in buttons:
                if btn.is_displayed() and btn.is_enabled():
                    # Use JavaScript click to avoid interception
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(2)
                    return True, "已点击「保存」按钮"
        
        # Try other common save button selectors
        save_selectors = [
            "button.ant-btn-primary",
            ".ant-modal-footer button.ant-btn-primary",
            ".modal-footer button.btn-primary",
            "button[type='submit']",
        ]
        
        for selector in save_selectors:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, selector)
                for btn in btns:
                    if btn.is_displayed() and btn.is_enabled():
                        btn_text = btn.text.strip().lower()
                        if "保存" in btn_text or "save" in btn_text or "确定" in btn_text or "ok" in btn_text or not btn_text:
                            driver.execute_script("arguments[0].click();", btn)
                            time.sleep(2)
                            return True, "已点击「保存」按钮"
            except Exception:
                continue
        
        return False, "未找到「保存」按钮"
    except Exception as e:
        return False, f"点击保存按钮失败: {str(e)}"


def _search_department_in_tree(driver, department_name):
    """Search for the newly created department in the left organization tree."""
    try:
        time.sleep(2)
        
        # Strategy: Find search input with placeholder="请输入关键字"
        search_input = None
        
        # Look for the specific search input in the organization tree
        search_selectors = [
            "input.ui-input[placeholder='请输入关键字']",
            "input[placeholder='请输入关键字']",
        ]
        
        for selector in search_selectors:
            try:
                inputs = driver.find_elements(By.CSS_SELECTOR, selector)
                for inp in inputs:
                    if inp.is_displayed() and not inp.get_attribute('readonly'):
                        # Check if this input is in a left-side container (x < screen width / 2)
                        rect = inp.rect
                        if rect['x'] < driver.get_window_size()['width'] / 2:
                            search_input = inp
                            break
                if search_input:
                    break
            except Exception:
                continue
        
        if not search_input:
            return False, "未找到组织树搜索框"
        
        # Clear and type the department name
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        
        ActionChains(driver).click(search_input).perform()
        time.sleep(0.3)
        
        try:
            search_input.clear()
        except Exception:
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.BACKSPACE)
        
        search_input.send_keys(department_name)
        time.sleep(1)
        
        # Trigger search
        search_input.send_keys(Keys.ENTER)
        time.sleep(2)
        
        # Check if the department appears in the tree
        # Look for elements containing the department name
        tree_items = driver.find_elements(By.CSS_SELECTOR, ".tree-node, .org-node, [role='treeitem'], .ant-tree-treenode, .department-node")
        
        found = False
        for item in tree_items:
            if item.is_displayed():
                item_text = item.text.strip()
                if department_name in item_text:
                    found = True
                    break
        
        # Also check body text as fallback
        if not found:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if department_name in body_text:
                found = True
        
        if found:
            return True, f"✅ 在组织树中找到部门: {department_name}"
        else:
            return False, f"❌ 在组织树中未找到部门: {department_name}"
            
    except Exception as e:
        return False, f"搜索部门失败: {str(e)}"


def create_new_department(driver: webdriver.Chrome, department_name: str = None) -> str:
    """
    Create a new department in eTeams.
    
    Steps:
    1. Ensure we're on the Org. Structure page (Division Info tab is already visible)
    2. Click New Department button
    3. Fill in department name (default: xuyingtest + timestamp)
    4. Click Save
    
    Args:
        driver: Selenium WebDriver instance
        department_name: Name for the new department (optional, defaults to xuyingtest + timestamp)
    
    Returns:
        Result message indicating success or failure
    """
    try:
        # Generate department name if not provided
        if not department_name:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            department_name = f"{DEFAULT_DEPARTMENT_PREFIX}{timestamp}"
        
        result_messages = []
        
        # Step 1: Verify we're on Org. Structure page
        current_url = driver.current_url
        if "/hrm/orgsetting/departmentSetting" not in current_url:
            return "错误: 当前不在组织架构设置页面，请先导航到该页面"
        
        result_messages.append("已在组织架构设置页面（分部信息页签已展开）")
        
        # Step 2: Click New Department button directly (no need to switch tabs)
        success, msg = _click_new_department_button(driver)
        result_messages.append(msg)
        if not success:
            return " | ".join(result_messages)
        
        # Step 3: Fill in department name
        success, msg, _, _, _ = _fill_department_form(driver, department_name)
        result_messages.append(msg)
        if not success:
            return " | ".join(result_messages)
        
        # Step 4: Click Save button
        success, msg = _click_save_button(driver)
        result_messages.append(msg)
        if not success:
            return " | ".join(result_messages)
        
        # Wait a bit for the save operation to complete
        time.sleep(2)
        
        result_messages.append(f"✅ 部门创建成功: {department_name}")
        
        # Step 5: Search for the department in the organization tree to verify
        search_success, search_msg = _search_department_in_tree(driver, department_name)
        result_messages.append(search_msg)
        
        if not search_success:
            return " | ".join(result_messages) + " | ⚠️ 警告: 未能在组织树中验证部门存在"
        
        return " | ".join(result_messages)
        
    except Exception as e:
        return f"❌ 创建部门失败: {str(e)}"


def open_browser_and_login():
    """Open browser and login to eTeams."""
    # Import here to avoid creating a selenium_tools -> create_new_department
    # circular import when this module is used as a helper by other tools.
    import selenium_tools

    # Open browser
    print("打开浏览器...")
    result = selenium_tools._open_eteams_login_page()
    print(f"浏览器打开结果: {result}\n")

    # Get driver
    driver = selenium_tools._driver
    if driver is None:
        raise Exception("浏览器未打开")

    # Set language to Chinese
    print("设置语言为简体中文...")
    selenium_tools._set_language_to_simplified_chinese(driver)

    # Fill login fields
    print("填写登录信息...")
    selenium_tools._fill_login_fields(driver, "13636409628", "xuying@0825")

    # Accept privacy terms
    print("接受隐私条款...")
    selenium_tools._accept_privacy_terms_if_needed(driver)

    # Click login button
    print("点击登录按钮...")
    selenium_tools._click_login_button(driver)

    # Wait for login result
    print("等待登录完成...")
    time.sleep(5)

    return driver


def test_create_department():
    """Test creating a new department."""
    driver = None
    try:
        # Step 1: Open browser and login
        print("=" * 60)
        print("步骤 1: 打开浏览器并登录...")
        print("=" * 60)
        driver = open_browser_and_login()
        print(f"当前 URL: {driver.current_url}\n")

        # Step 2: Navigate to Org. Structure page
        print("=" * 60)
        print("步骤 2: 导航到组织架构设置页面...")
        print("=" * 60)

        # Navigate to org structure page using the correct URL
        org_structure_url = "https://weapp.eteams.cn/hrm/orgsetting/departmentSetting"
        print(f"正在导航到: {org_structure_url}")
        driver.get(org_structure_url)

        # Wait for page to load
        print("等待页面加载...")
        time.sleep(5)
        print(f"当前 URL: {driver.current_url}")
        print(f"页面标题: {driver.title}\n")

        # Step 3: Create new department
        print("=" * 60)
        print("步骤 3: 创建新部门...")
        print("=" * 60)
        result = create_new_department(driver)
        print(f"结果: {result}\n")

        print("=" * 60)
        print("✅ 测试完成!")
        print("=" * 60)
        return True

    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback

        traceback.print_exc()
        return False
    finally:
        # Keep browser open for inspection
        if driver:
            print("\n浏览器保持打开状态，请手动关闭")


__all__ = [
    "create_new_department",
    "open_browser_and_login",
    "test_create_department",
]


if __name__ == "__main__":
    success = test_create_department()
    sys.exit(0 if success else 1)
