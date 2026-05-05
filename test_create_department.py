"""
Test script for creating a new department in eTeams.
"""

import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from create_new_department import create_new_department

# Import internal functions from selenium_tools
import selenium_tools


def open_browser_and_login():
    """Open browser and login to eTeams."""
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


if __name__ == "__main__":
    success = test_create_department()
    sys.exit(0 if success else 1)
