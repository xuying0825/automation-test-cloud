"""
Selenium tools for testing the eTeams Passport login page
(https://passport.eteams.cn/).

Each tool is a Python function decorated with @function_tool for use with the
openai-agents SDK.
"""

import logging
import os
import secrets
import time
from datetime import datetime
from agents import function_tool
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Setup logger
logger = logging.getLogger(__name__)
from create_public_group import create_public_group as _standalone_create_public_group
from create_new_person import create_new_person as _standalone_create_new_person
from create_new_department import create_new_department as _standalone_create_new_department
from edit_person import (
    edit_person_employee_no_and_hire_date
    as _standalone_edit_person_employee_no_and_hire_date,
)

from select_org_structure import select_org_structure as _standalone_select_org_structure

# Global browser instance shared across tool calls within a session
_driver: webdriver.Chrome | None = None

TARGET_URL = "https://passport.eteams.cn/"
ORG_STRUCTURE_PATH = "/hrm/orgsetting/departmentSetting"
DEFAULT_ACCOUNT = os.environ.get("ETEAMS_ACCOUNT", "13636409628")
_DEFAULT_PASSWORD = os.environ.get("ETEAMS_PASSWORD", "xuying@0825")
TEST_GROUP_NAME_OVERRIDE = os.environ.get("ETEAMS_TEST_GROUP_NAME", "").strip()
DEFAULT_TEST_GROUP_PREFIX = os.environ.get("ETEAMS_TEST_GROUP_PREFIX", "xuyingtest")
DEFAULT_CLOSE_DELAY_SECONDS = int(os.environ.get("ETEAMS_CLOSE_DELAY_SECONDS", "5"))
DEFAULT_BASIC_LOGIN_CLOSE_DELAY_SECONDS = int(
    os.environ.get("ETEAMS_BASIC_LOGIN_CLOSE_DELAY_SECONDS", "5")
)
DEFAULT_IMPLICIT_WAIT_SECONDS = 5
LOGIN_RESULT_TIMEOUT_SECONDS = int(os.environ.get("ETEAMS_LOGIN_RESULT_TIMEOUT_SECONDS", "30"))


def _mask_account(account: str) -> str:
    """Mask an account/phone number for tool outputs."""
    if not account:
        return ""
    if len(account) >= 7:
        return f"{account[:3]}****{account[-4:]}"
    return "***"


def _generate_test_group_name() -> str:
    """Generate the public group name for this run."""
    if TEST_GROUP_NAME_OVERRIDE:
        return TEST_GROUP_NAME_OVERRIDE

    # eTeams accepts simple text names; keep the prefix recognizable while
    # adding timestamp + random suffix so repeated self-tests create a fresh
    # group instead of colliding with prior runs.
    timestamp = datetime.now().strftime("%m%d%H%M%S")
    suffix = secrets.token_hex(2)
    return f"{DEFAULT_TEST_GROUP_PREFIX}{timestamp}{suffix}"


def _get_windows_descendant_process_ids(root_pid: int | None) -> set[int]:
    """Return descendant process IDs for a Windows process, best effort."""
    if os.name != "nt" or not root_pid:
        return set()

    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return set()

    TH32CS_SNAPPROCESS = 0x00000002
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.windll.kernel32
    try:
        kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
        kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        kernel32.Process32FirstW.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(PROCESSENTRY32W),
        ]
        kernel32.Process32FirstW.restype = wintypes.BOOL
        kernel32.Process32NextW.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(PROCESSENTRY32W),
        ]
        kernel32.Process32NextW.restype = wintypes.BOOL
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL
    except Exception:
        pass

    snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snapshot or snapshot == INVALID_HANDLE_VALUE:
        return set()

    parent_to_children: dict[int, set[int]] = {}
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        has_entry = kernel32.Process32FirstW(snapshot, ctypes.byref(entry))
        while has_entry:
            pid = int(entry.th32ProcessID)
            parent_pid = int(entry.th32ParentProcessID)
            parent_to_children.setdefault(parent_pid, set()).add(pid)
            has_entry = kernel32.Process32NextW(snapshot, ctypes.byref(entry))
    except Exception:
        return set()
    finally:
        try:
            kernel32.CloseHandle(snapshot)
        except Exception:
            pass

    descendants: set[int] = set()
    pending = list(parent_to_children.get(int(root_pid), set()))
    while pending:
        pid = pending.pop()
        if pid in descendants:
            continue
        descendants.add(pid)
        pending.extend(parent_to_children.get(pid, set()))
    return descendants


def _find_chrome_window_handle_on_windows(driver: webdriver.Chrome) -> int | None:
    """Find the native HWND for Selenium's Chrome window on Windows."""
    if os.name != "nt":
        return None

    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None

    try:
        driver_title = (driver.title or "").strip()
    except Exception:
        driver_title = ""

    chrome_driver_pid = None
    try:
        chrome_driver_pid = int(driver.service.process.pid)
    except Exception:
        chrome_driver_pid = None
    selenium_child_pids = _get_windows_descendant_process_ids(chrome_driver_pid)

    user32 = ctypes.windll.user32

    enum_windows_proc = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HWND,
        wintypes.LPARAM,
    )

    try:
        user32.EnumWindows.argtypes = [enum_windows_proc, wintypes.LPARAM]
        user32.EnumWindows.restype = wintypes.BOOL
        user32.IsWindowVisible.argtypes = [wintypes.HWND]
        user32.IsWindowVisible.restype = wintypes.BOOL
        user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
        user32.GetWindowTextLengthW.restype = ctypes.c_int
        user32.GetWindowTextW.argtypes = [
            wintypes.HWND,
            wintypes.LPWSTR,
            ctypes.c_int,
        ]
        user32.GetWindowTextW.restype = ctypes.c_int
        user32.GetWindowThreadProcessId.argtypes = [
            wintypes.HWND,
            ctypes.POINTER(wintypes.DWORD),
        ]
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        user32.GetClassNameW.argtypes = [
            wintypes.HWND,
            wintypes.LPWSTR,
            ctypes.c_int,
        ]
        user32.GetClassNameW.restype = ctypes.c_int
    except Exception:
        # ctypes signatures are a safety improvement, but failure should not
        # prevent the best-effort window search.
        pass

    def _window_text(hwnd: int) -> str:
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value.strip()

    def _class_name(hwnd: int) -> str:
        buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buffer, len(buffer))
        return buffer.value.strip()

    title_key = driver_title.casefold()

    candidates: list[tuple[int, int, str]] = []

    def _enum_window(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        class_name = _class_name(hwnd)
        if not class_name.startswith("Chrome_WidgetWin"):
            return True

        window_title = _window_text(hwnd)
        pid_value = wintypes.DWORD(0)
        try:
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_value))
        except Exception:
            pid_value = wintypes.DWORD(0)
        window_pid = int(pid_value.value or 0)
        is_selenium_child = bool(
            selenium_child_pids and window_pid in selenium_child_pids
        )

        if not window_title and not is_selenium_child:
            return True

        window_key = window_title.casefold()
        score = 10
        if is_selenium_child:
            score += 250
        if class_name == "Chrome_WidgetWin_1":
            score += 5
        if title_key and title_key in window_key:
            score += 100
        if title_key == "data:," and "data:," in window_key:
            score += 100
        if "chrome" in window_key:
            score += 1

        candidates.append((score, int(hwnd), window_title))
        return True

    callback = enum_windows_proc(_enum_window)
    try:
        user32.EnumWindows(callback, 0)
    except Exception:
        return None

    if not candidates:
        return None

    # Prefer the window whose native title contains Selenium's current page
    # title. If the page is still the initial data:, page, that title is usually
    # unique enough to avoid focusing the user's chat browser by mistake.
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def _activate_window_handle_on_windows(hwnd: int) -> bool:
    """Best-effort native foreground activation for a Windows HWND."""
    if os.name != "nt" or not hwnd:
        return False

    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return False

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    SW_RESTORE = 9
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    SWP_NOMOVE = 0x0002
    SWP_NOSIZE = 0x0001
    SWP_SHOWWINDOW = 0x0040

    try:
        user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.ShowWindow.restype = wintypes.BOOL
        user32.SetWindowPos.argtypes = [
            wintypes.HWND,
            wintypes.HWND,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_uint,
        ]
        user32.SetWindowPos.restype = wintypes.BOOL
        user32.BringWindowToTop.argtypes = [wintypes.HWND]
        user32.BringWindowToTop.restype = wintypes.BOOL
        user32.SetForegroundWindow.argtypes = [wintypes.HWND]
        user32.SetForegroundWindow.restype = wintypes.BOOL
        user32.SetActiveWindow.argtypes = [wintypes.HWND]
        user32.SetActiveWindow.restype = wintypes.HWND
        user32.GetForegroundWindow.restype = wintypes.HWND
        user32.GetWindowThreadProcessId.argtypes = [
            wintypes.HWND,
            ctypes.POINTER(wintypes.DWORD),
        ]
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        user32.AttachThreadInput.argtypes = [
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.BOOL,
        ]
        user32.AttachThreadInput.restype = wintypes.BOOL
        kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    except Exception:
        pass

    current_thread = kernel32.GetCurrentThreadId()
    target_thread = user32.GetWindowThreadProcessId(hwnd, None)
    foreground_hwnd = user32.GetForegroundWindow()
    foreground_thread = (
        user32.GetWindowThreadProcessId(foreground_hwnd, None)
        if foreground_hwnd
        else 0
    )

    attached_target = False
    attached_foreground = False
    try:
        if target_thread and target_thread != current_thread:
            attached_target = bool(user32.AttachThreadInput(current_thread, target_thread, True))
        if foreground_thread and foreground_thread != current_thread:
            attached_foreground = bool(
                user32.AttachThreadInput(current_thread, foreground_thread, True)
            )

        user32.ShowWindow(hwnd, SW_RESTORE)
        # The topmost flip works around Windows' focus-stealing prevention when
        # Selenium is triggered from a background Flask/server thread.
        user32.SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
        )
        user32.SetWindowPos(
            hwnd,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
        )
        user32.BringWindowToTop(hwnd)
        user32.SetActiveWindow(hwnd)
        return bool(user32.SetForegroundWindow(hwnd))
    except Exception:
        return False
    finally:
        if attached_foreground:
            user32.AttachThreadInput(current_thread, foreground_thread, False)
        if attached_target:
            user32.AttachThreadInput(current_thread, target_thread, False)


def _bring_driver_window_to_front(driver: webdriver.Chrome) -> None:
    """Bring the Selenium-controlled browser window to the foreground.

    Selenium can create Chrome successfully while Windows keeps the previously
    active browser (for example the Flask chat UI) in front. The WebDriver
    focus calls only affect the DOM/tab, so on Windows we additionally activate
    the native Chrome window with Win32 APIs.
    """
    try:
        driver.switch_to.window(driver.current_window_handle)
    except Exception:
        pass

    try:
        driver.execute_cdp_cmd("Page.bringToFront", {})
    except Exception:
        pass

    try:
        driver.execute_script("window.focus();")
    except Exception:
        pass

    if os.name == "nt":
        hwnd = _find_chrome_window_handle_on_windows(driver)
        if hwnd:
            activated = _activate_window_handle_on_windows(hwnd)
            logger.info(
                "Windows Chrome 窗口激活%s，hwnd=%s",
                "成功" if activated else "未确认成功",
                hwnd,
            )
        else:
            logger.warning("Windows 下未找到 Selenium Chrome 原生窗口句柄。")


def _get_driver() -> webdriver.Chrome:
    """Get or create the Chrome driver and keep its window in front."""
    global _driver
    if _driver is None:
        logger.info("初始化 Chrome 浏览器...")
        options = webdriver.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1440,1000")
        # On Windows the Flask chat page can stay in front of the Selenium
        # browser. Starting maximized plus the explicit focus logic below makes
        # the Chrome window much easier to notice when a tool is triggered from
        # the background agent thread.
        options.add_argument("--start-maximized")
        # Suppress Chrome's browser-level "Save password?" bubble. Selenium
        # cannot reliably locate/click that UI because it is not part of the
        # web page DOM, so the safest automation behavior is to never save.
        options.add_experimental_option("detach", True)
        options.add_experimental_option(
            "prefs",
            {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "profile.password_manager_leak_detection": False,
            },
        )
        # Try to use local ChromeDriver first, fallback to WebDriver Manager
        local_driver_paths = [
            r"C:\WebDriver\chromedriver.exe",  # Windows
            "/usr/local/bin/chromedriver",      # macOS/Linux
        ]
        
        service = None
        for driver_path in local_driver_paths:
            if os.path.exists(driver_path):
                logger.info(f"使用本地 ChromeDriver: {driver_path}")
                service = Service(driver_path)
                break
        
        if service is None:
            logger.info("未找到本地 ChromeDriver，使用 WebDriver Manager 自动下载")
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
        
        _driver = webdriver.Chrome(service=service, options=options)
        _driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)
        try:
            if os.name == "nt":
                _driver.set_window_position(0, 0)
                _driver.maximize_window()
            else:
                _driver.set_window_size(1440, 1000)
        except Exception as exc:
            logger.warning("设置 Chrome 窗口尺寸/位置失败：%s", exc)
        logger.info("Chrome 浏览器初始化完成，session_id=%s", _driver.session_id)

    _bring_driver_window_to_front(_driver)
    return _driver


def _close_driver() -> bool:
    """Close the shared Chrome driver if it is open."""
    global _driver
    if _driver is None:
        return False
    _driver.quit()
    _driver = None
    return True


def _save_screenshot(driver: webdriver.Chrome, label: str) -> str:
    """Save a screenshot under screenshots/ and return a relative path message."""
    screenshots_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "screenshots"
    )
    os.makedirs(screenshots_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = "".join(c if c.isalnum() or c in "-_" else "_" for c in label)
    filename = f"{timestamp}_{safe_label}.png"
    filepath = os.path.join(screenshots_dir, filename)

    driver.save_screenshot(filepath)
    return f"📸 截图已保存：screenshots/{filename}"


def _wait_for_login_form(driver: webdriver.Chrome, timeout: int = 20):
    """Wait until the eTeams account/password fields are visible."""

    def _find_fields(drv: webdriver.Chrome):
        fields = []
        for element in drv.find_elements(By.CSS_SELECTOR, "input.ui-input"):
            try:
                if element.is_displayed():
                    fields.append(element)
            except StaleElementReferenceException:
                return False
        return fields[:2] if len(fields) >= 2 else False

    return WebDriverWait(driver, timeout).until(_find_fields)


def _login_form_is_present(
    driver: webdriver.Chrome,
    *,
    fast: bool = False,
) -> bool:
    if fast:
        driver.implicitly_wait(0)
    try:
        fields = [
            element
            for element in driver.find_elements(By.CSS_SELECTOR, "input.ui-input")
            if element.is_displayed()
        ]
        if len(fields) < 2:
            return False

        placeholders = [
            (field.get_attribute("placeholder") or "").strip().lower()
            for field in fields
        ]
        field_types = [
            (field.get_attribute("type") or "").strip().lower()
            for field in fields
        ]
        has_account_field = any(
            token in placeholder
            for placeholder in placeholders
            for token in ("账号", "帐号", "手机", "account", "username", "phone")
        )
        has_password_field = any(
            token in placeholder
            for placeholder in placeholders
            for token in ("密码", "password")
        ) or "password" in field_types
        if has_account_field and has_password_field:
            return True

        body_text = driver.find_element(By.TAG_NAME, "body").text
        return (
            len(fields) >= 2
            and ("登录" in body_text or "Login" in body_text)
            and ("密码" in body_text or "Password" in body_text)
        )
    except Exception:
        return False
    finally:
        if fast:
            driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)


def _is_simplified_chinese(driver: webdriver.Chrome) -> bool:
    """Return whether the login page is currently displayed in 简体中文."""
    try:
        fields = [
            element
            for element in driver.find_elements(By.CSS_SELECTOR, "input.ui-input")
            if element.is_displayed()
        ]
        placeholders = [field.get_attribute("placeholder") for field in fields]
        body_text = driver.find_element(By.TAG_NAME, "body").text
        return (
            "账号" in placeholders
            or "选择语言" in body_text
            or driver.title.strip() == "登录"
        )
    except Exception:
        return False


def _wait_for_visible_xpath(
    driver: webdriver.Chrome,
    xpath: str,
    timeout: int = 10,
):
    """Wait for the first displayed element matching an XPath."""

    def _find_element(drv: webdriver.Chrome):
        for element in drv.find_elements(By.XPATH, xpath):
            try:
                if element.is_displayed():
                    return element
            except StaleElementReferenceException:
                return False
        return False

    return WebDriverWait(driver, timeout).until(_find_element)


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


def _set_language_to_simplified_chinese(driver: webdriver.Chrome) -> str:
    """Switch the eTeams login page language to 简体中文."""
    _wait_for_login_form(driver)

    if _is_simplified_chinese(driver):
        return "语言已是简体中文。"

    try:
        trigger = _wait_for_visible_xpath(
            driver,
            (
                "//*[contains(@class, 'weapp-passport-loginCom-i18n')]"
                " | //*[normalize-space()='Select Language' or normalize-space()='选择语言']"
            ),
        )
        _safe_click(driver, trigger)

        simplified_chinese = _wait_for_visible_xpath(
            driver,
            (
                "//div[contains(@class, 'weapp-passport-i18n-langs-item')"
                " and normalize-space()='简体中文']"
            ),
        )
        _safe_click(driver, simplified_chinese)

        WebDriverWait(driver, 20).until(lambda drv: _is_simplified_chinese(drv))
        _wait_for_login_form(driver)
        return "已通过页面语言菜单切换为简体中文。"
    except Exception:
        # Fallback for cases where the popup is blocked or slow: eTeams stores
        # the selected language in localStorage.langType.
        driver.execute_script(
            """
            localStorage.setItem('langType', 'zh_CN');
            sessionStorage.setItem('PASSPORT_LANGTYPE_CHANGED', '1');
            """
        )
        driver.refresh()
        WebDriverWait(driver, 20).until(lambda drv: _is_simplified_chinese(drv))
        _wait_for_login_form(driver)
        return "已通过本地语言配置切换为简体中文。"


def _fill_login_fields(
    driver: webdriver.Chrome,
    username: str,
    password: str,
) -> None:
    account_field, password_field = _wait_for_login_form(driver)

    _safe_click(driver, account_field)
    account_field.clear()
    account_field.send_keys(username)
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
        account_field,
        username,
    )

    _safe_click(driver, password_field)
    password_field.clear()
    password_field.send_keys(password)
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
        password_field,
        password,
    )


def _accept_privacy_terms_if_needed(driver: webdriver.Chrome) -> str:
    """Tick only the privacy checkbox, without opening the policy links."""
    try:
        checkbox = driver.find_element(By.CSS_SELECTOR, "input.ui-checkbox-input")
    except NoSuchElementException:
        return "未找到隐私协议勾选框，跳过勾选。"

    try:
        if driver.execute_script("return arguments[0].checked === true;", checkbox):
            return "「我已阅读并同意」已勾选。"
    except Exception:
        pass

    # Do not click the whole label/wrapper: it contains the "用户条款" and
    # "隐私政策" hyperlinks, and clicking the wrapper can accidentally open
    # https://eteams.cn/services?... instead of just checking the box.
    driver.execute_script(
        """
        const input = arguments[0];
        input.scrollIntoView({block: 'center', inline: 'center'});
        if (input.checked !== true) {
          input.click();
        }
        """,
        checkbox,
    )

    try:
        WebDriverWait(driver, 5).until(
            lambda drv: drv.execute_script(
                "return arguments[0].checked === true;",
                checkbox,
            )
        )
        return "已勾选「我已阅读并同意」。"
    except TimeoutException:
        return "已尝试勾选「我已阅读并同意」。"


def _click_login_button(driver: webdriver.Chrome) -> str:
    """Click the real 登录 button and verify that submission starts."""
    login_button_xpath = (
        "//*[(self::button or @role='button' or contains(@class, 'button')"
        " or contains(@class, 'Button'))"
        " and (normalize-space()='登录' or normalize-space()='Login'"
        " or .//*[normalize-space()='登录' or normalize-space()='Login'])]"
    )

    def _find_clickable_button(drv: webdriver.Chrome):
        for element in drv.find_elements(By.XPATH, login_button_xpath):
            try:
                if element.is_displayed() and element.is_enabled():
                    return element
            except StaleElementReferenceException:
                return False
        return False

    def _button_is_loading_or_disabled(drv: webdriver.Chrome) -> bool:
        try:
            return bool(
                drv.execute_script(
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
                        && rect.height > 0;
                    };
                    for (const el of Array.from(document.querySelectorAll('button, [role="button"], [class*="button"]'))) {
                      if (!visible(el)) continue;
                      const text = normalize(el.innerText || el.textContent);
                      if (!['登录', 'Login'].includes(text)) continue;
                      const attrs = [
                        el.className,
                        el.getAttribute('class'),
                        el.getAttribute('aria-busy'),
                        el.getAttribute('disabled')
                      ].map(normalize).join(' ').toLowerCase();
                      return Boolean(el.disabled)
                        || el.getAttribute('aria-disabled') === 'true'
                        || /loading|disabled|submitting|pending|加载/.test(attrs);
                    }
                    return false;
                    """
                )
            )
        except Exception:
            return False

    def _submission_started(start_url: str) -> bool:
        current_url = driver.current_url.rstrip("/")
        if current_url != start_url:
            return True
        if not _login_form_is_present(driver, fast=True):
            return True
        if _button_is_loading_or_disabled(driver):
            return True
        if _collect_visible_messages(driver, fast=True):
            return True
        return False

    def _wait_for_submission_start(start_url: str, timeout: float = 3.0) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if _submission_started(start_url):
                return True
            time.sleep(0.2)
        return _submission_started(start_url)

    def _js_click_button(element) -> None:
        driver.execute_script(
            """
            const el = arguments[0];
            el.scrollIntoView({block: 'center', inline: 'center'});
            const rect = el.getBoundingClientRect();
            const x = Math.floor(rect.left + rect.width / 2);
            const y = Math.floor(rect.top + rect.height / 2);
            const target = document.elementFromPoint(x, y) || el;
            const opts = {
              bubbles: true,
              cancelable: true,
              view: window,
              clientX: x,
              clientY: y,
              button: 0,
              buttons: 1
            };
            for (const type of ['pointerover', 'pointermove', 'pointerdown']) {
              if (window.PointerEvent) {
                target.dispatchEvent(new PointerEvent(type, opts));
              }
            }
            for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
              target.dispatchEvent(new MouseEvent(type, opts));
            }
            if (window.PointerEvent) {
              target.dispatchEvent(new PointerEvent('pointerup', opts));
            }
            if (typeof el.click === 'function') {
              el.click();
            }
            """,
            element,
        )

    start_url = driver.current_url.rstrip("/")
    errors: list[str] = []
    strategies = (
        "actionchains-click",
        "webdriver-click",
        "javascript-pointer-click",
        "password-enter",
    )

    for attempt, strategy in enumerate(strategies, start=1):
        try:
            if strategy == "password-enter":
                _, password_field = _wait_for_login_form(driver, timeout=5)
                _safe_click(driver, password_field)
                password_field.send_keys(Keys.ENTER)
            else:
                login_button = WebDriverWait(driver, 10).until(_find_clickable_button)
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    login_button,
                )
                if strategy == "webdriver-click":
                    login_button.click()
                elif strategy == "actionchains-click":
                    ActionChains(driver).move_to_element(login_button).pause(0.15).click(
                        login_button
                    ).perform()
                else:
                    _js_click_button(login_button)

            if _wait_for_submission_start(start_url):
                return f"已触发登录提交（第 {attempt} 次，方式：{strategy}）。"
        except Exception as exc:
            errors.append(f"{strategy}: {str(exc)}")

        time.sleep(0.4)

    if errors:
        return "已多次尝试点击登录，但页面暂未发生变化；点击异常：" + "；".join(errors)
    return "已多次尝试点击登录，但页面暂未发生变化。"


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


def _decline_save_password_prompt_if_possible(driver: webdriver.Chrome) -> str:
    """
    Ensure browser password saving is declined.

    Chrome's "Save password?" prompt is browser chrome, not page DOM, so it is
    not reliably clickable through Selenium. We disable it via Chrome prefs when
    creating the driver; ESC is an extra best-effort fallback if a prompt still
    appears.
    """
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        return "已配置不保存密码；如出现 Save password 提示，已尝试选择 No Thanks/取消。"
    except Exception:
        return "已配置不保存密码。"


def _find_top_left_logged_in_user(driver: webdriver.Chrome) -> str:
    """
    Return visible text that looks like the logged-in user in the upper-left
    area of the post-login page.
    """
    if _login_form_is_present(driver, fast=True):
        return ""

    driver.implicitly_wait(0)
    try:
        candidates = driver.execute_script(
            """
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const ignoredTags = new Set([
              'SCRIPT', 'STYLE', 'NOSCRIPT', 'META', 'LINK', 'SVG', 'PATH'
            ]);
            const ignoredTexts = new Set([
              'eteams', 'eTeams', '首页', '工作台', '消息', '通讯录', '应用',
              '管理', '登录', '账号', '密码', '选择语言', 'Login',
              'Select Language'
            ]);
            const isVisible = (el) => {
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
            const likelyUserAttrs = (el) => {
              const attrs = [
                el.id,
                el.className,
                el.getAttribute('title'),
                el.getAttribute('aria-label'),
                el.getAttribute('data-name'),
                el.getAttribute('data-user-name')
              ].map(normalize).join(' ').toLowerCase();
              return /user|account|avatar|profile|member|staff|employee|person|username|nickname|name|login|用户|账号|姓名|成员|头像/.test(attrs);
            };

            const result = [];
            for (const el of Array.from(document.body.querySelectorAll('*'))) {
              if (ignoredTags.has(el.tagName) || !isVisible(el)) {
                continue;
              }
              const rect = el.getBoundingClientRect();
              if (rect.left > 520 || rect.top > 220) {
                continue;
              }

              const text = normalize(
                el.innerText
                || el.textContent
                || el.getAttribute('title')
                || el.getAttribute('aria-label')
              );
              if (!text || text.length < 2 || text.length > 80 || ignoredTexts.has(text)) {
                continue;
              }

              const score = (likelyUserAttrs(el) ? 100 : 0)
                + Math.max(0, 220 - rect.top) / 10
                + Math.max(0, 520 - rect.left) / 20
                - text.length / 100;
              result.push({ text, score, top: rect.top, left: rect.left });
            }

            result.sort((a, b) =>
              b.score - a.score || a.top - b.top || a.left - b.left
            );
            return result.slice(0, 5);
            """
        )
    except Exception:
        return ""
    finally:
        driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)

    for candidate in candidates or []:
        text = str(candidate.get("text", "")).strip()
        if text:
            return text
    return ""


def _wait_for_login_result(
    driver: webdriver.Chrome,
    timeout: int = LOGIN_RESULT_TIMEOUT_SECONDS,
) -> dict:
    """
    Wait until login succeeds/fails/gets stuck.

    Success is confirmed only when the logged-in user text is visible in the
    upper-left area, per the test requirement.
    """
    start_url = driver.current_url.rstrip("/")
    deadline = time.monotonic() + timeout
    last_messages: list[str] = []

    while time.monotonic() < deadline:
        top_left_user = _find_top_left_logged_in_user(driver)
        if top_left_user:
            return {
                "status": "success",
                "top_left_user": top_left_user,
                "messages": [],
            }

        last_messages = _collect_visible_messages(driver, fast=True)
        if last_messages and _login_form_is_present(driver, fast=True):
            return {
                "status": "failed",
                "top_left_user": "",
                "messages": last_messages,
            }

        time.sleep(0.5)

    top_left_user = _find_top_left_logged_in_user(driver)
    if top_left_user:
        return {
            "status": "success",
            "top_left_user": top_left_user,
            "messages": [],
        }

    messages = _collect_visible_messages(driver, fast=True) or last_messages
    if messages:
        return {
            "status": "failed",
            "top_left_user": "",
            "messages": messages,
        }

    login_form_present = _login_form_is_present(driver, fast=True)
    current_url = driver.current_url.rstrip("/")
    return {
        "status": "unknown",
        "top_left_user": "",
        "messages": [],
        "page_changed": current_url != start_url or not login_form_present,
    }


def _find_visible_org_structure_menu_item(driver: webdriver.Chrome):
    """Find a visible Org.Structure menu item after the top-right menu opens."""
    driver.implicitly_wait(0)
    try:
        candidates = driver.execute_script(
            """
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const visible = (el) => {
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
            const labelPatterns = [
              'Org.Structure',
              'Org Structure',
              'Organization Structure',
              '组织架构',
              '组织结构'
            ];
            const itemSelector = [
              'a',
              'button',
              'li',
              '[role="menuitem"]',
              '[class*="item"]',
              '[class*="Item"]',
              'div',
              'span'
            ].join(',');
            const result = [];
            for (const el of Array.from(document.body.querySelectorAll(itemSelector))) {
              if (!visible(el)) {
                continue;
              }
              const text = normalize(el.innerText || el.textContent);
              if (!text || !labelPatterns.some((label) => text.includes(label))) {
                continue;
              }

              const clickable = el.closest(
                'a,button,li,[role="menuitem"],[class*="item"],[class*="Item"]'
              ) || el;
              if (!visible(clickable)) {
                continue;
              }

              const clickableText = normalize(clickable.innerText || clickable.textContent);
              const rect = clickable.getBoundingClientRect();
              const exact = labelPatterns.includes(clickableText) ? 1 : 0;
              const area = rect.width * rect.height;
              result.push({ element: clickable, exact, length: clickableText.length, area });
            }
            result.sort((a, b) =>
              b.exact - a.exact || a.length - b.length || a.area - b.area
            );
            return result.slice(0, 5).map((item) => item.element);
            """
        )
    finally:
        driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)

    for element in candidates or []:
        try:
            if element.is_displayed() and element.is_enabled():
                return element
        except StaleElementReferenceException:
            continue
    return False


def _get_top_right_dropdown_candidates(driver: webdriver.Chrome):
    """
    Return likely clickable candidates for the logged-in user's top-right menu.

    The exact eTeams DOM can vary, so this scores visible clickable elements in
    the top-right region using common user/menu/avatar/dropdown attributes.
    """
    driver.implicitly_wait(0)
    try:
        return driver.execute_script(
            """
            const normalize = (value) => String(value || '')
              .replace(/\\s+/g, ' ')
              .trim();
            const visible = (el) => {
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
            const clickableFor = (el) =>
              el.closest(
                'button,a,[role="button"],[aria-haspopup="true"],'
                + '[class*="dropdown"],[class*="Dropdown"],'
                + '[class*="menu"],[class*="Menu"],'
                + '[class*="avatar"],[class*="Avatar"],'
                + '[class*="user"],[class*="User"],'
                + '[class*="profile"],[class*="Profile"]'
              ) || el;

            const seen = new Set();
            const candidates = [];
            for (const raw of Array.from(document.body.querySelectorAll('*'))) {
              if (!visible(raw)) {
                continue;
              }
              const rect = raw.getBoundingClientRect();
              if (rect.left < window.innerWidth - 560 || rect.top > 180) {
                continue;
              }

              const el = clickableFor(raw);
              if (seen.has(el) || !visible(el)) {
                continue;
              }
              seen.add(el);

              const elRect = el.getBoundingClientRect();
              if (elRect.left < window.innerWidth - 620 || elRect.top > 210) {
                continue;
              }

              const attrs = [
                el.tagName,
                el.id,
                el.className,
                el.getAttribute('title'),
                el.getAttribute('aria-label'),
                el.getAttribute('aria-haspopup'),
                el.getAttribute('role'),
                normalize(el.innerText || el.textContent)
              ].map(normalize).join(' ').toLowerCase();
              const style = window.getComputedStyle(el);

              let score = 0;
              if (/user|account|avatar|profile|member|staff|employee|person|dropdown|menu|setting|more|用户|账号|姓名|成员|头像|菜单|更多|设置/.test(attrs)) {
                score += 120;
              }
              if (style.cursor === 'pointer') {
                score += 40;
              }
              if (['BUTTON', 'A'].includes(el.tagName)) {
                score += 30;
              }
              if (el.getAttribute('aria-haspopup') === 'true') {
                score += 50;
              }
              score += Math.max(0, elRect.left - (window.innerWidth - 620)) / 20;
              score += Math.max(0, 210 - elRect.top) / 10;

              candidates.push({ element: el, score, top: elRect.top, left: elRect.left });
            }

            candidates.sort((a, b) =>
              b.score - a.score || b.left - a.left || a.top - b.top
            );
            return candidates.slice(0, 12).map((item) => item.element);
            """
        )
    finally:
        driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT_SECONDS)


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
    The dropdown may have varying number of items depending on screen size and scroll position.
    We check for the presence of "组织架构设置" (Org. Structure) instead of validating exact item count.
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

    def _wait_for_eteams_dropdown_trigger(timeout: int = 5) -> bool:
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
            # Check if "组织架构设置" is present in the dropdown
            has_org_structure = any(
                "组织架构设置" in item or "Org. Structure" in item or "Org.Structure" in item
                for item in last_items
            )
            if len(last_items) >= 3 and has_org_structure:
                return last_items
            time.sleep(0.25)
        return last_items

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
            # Check if "组织架构设置" is present
            has_org_structure = any(
                "组织架构设置" in item or "Org. Structure" in item or "Org.Structure" in item
                for item in last_items
            )
            if len(last_items) >= 3 and has_org_structure:
                break

            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except Exception:
                pass
            time.sleep(0.5)

        # Verify that "组织架构设置" is present in the dropdown
        has_org_structure = any(
            "组织架构设置" in item or "Org. Structure" in item or "Org.Structure" in item
            for item in last_items
        )
        if not has_org_structure:
            return (
                "未能选择 Org. Structure：已找到右上角 eTeams 下拉菜单触发器，"
                f"但未找到「组织架构设置」菜单项。菜单项：{', '.join(last_items)}"
            )

        clicked_text = _click_org_structure_menu_item()
        if not clicked_text:
            return (
                "未能选择 Org. Structure：右上角 eTeams 下拉菜单已展开，"
                f"但未找到 Org. Structure/组织架构设置。菜单项：{', '.join(last_items)}"
            )

        WebDriverWait(driver, 12, poll_frequency=0.5).until(_org_structure_page_is_open)
        return (
            f"已找到右上角 eTeams 下拉菜单（{len(last_items)}个选项），并点击 {clicked_text}。"
        )
    except Exception as exc:
        return f"未能选择 Org. Structure：{str(exc)}"


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


def select_org_structure(driver: webdriver.Chrome | None = None) -> str:
    """
    Find and select Org. Structure from the real top-right eTeams dropdown.

    Preconditions:
    - The browser must already be logged in to eTeams.
    - The top-right eTeams logo/caret dropdown must be available in the current
      window size.

    This function deliberately does not use a fixed URL fallback and does not
    resize the browser window.
    """
    driver = driver or _get_driver()
    return _standalone_select_org_structure(driver)


def create_public_group(
    driver: webdriver.Chrome | None = None,
    group_name: str = "",
) -> str:
    """
    Open left-side 群组管理 in Org. Structure and create a public group.

    Preconditions:
    - Call ``select_org_structure(driver)`` first, or otherwise make sure the
      current page is already Org. Structure.

    Args:
        driver: Existing Selenium Chrome driver. If omitted, uses this module's
            shared driver.
        group_name: Group name to create. If omitted, a random name prefixed by
            ``xuyingtest`` is generated for this run.
    """
    driver = driver or _get_driver()
    return _standalone_create_public_group(driver, group_name=group_name)


def create_new_person(driver: webdriver.Chrome | None = None) -> str:
    """
    Open 组织维护 > 人力资源 in Org. Structure and create a new person.

    Preconditions:
    - The browser should already be logged in to eTeams.
    - The function will enter/verify Org. Structure before continuing.
    """
    driver = driver or _get_driver()
    return _standalone_create_new_person(driver)


def create_new_department(
    driver: webdriver.Chrome | None = None,
    department_name: str = "",
) -> str:
    """
    Open Org. Structure / 组织维护 / 分部信息 and create a new department.

    Preconditions:
    - The browser should already be logged in to eTeams.
    - The caller should have entered Org. Structure; the standalone helper
      contains the already-tested department creation steps.

    Args:
        driver: Existing Selenium Chrome driver. If omitted, uses this module's
            shared driver.
        department_name: Department name to create. If omitted, a random name
            prefixed by ``xuyingtest`` is generated.
    """
    driver = driver or _get_driver()
    return _standalone_create_new_department(
        driver,
        department_name=department_name or None,
    )


def edit_person(
    driver: webdriver.Chrome | None = None,
    person_name: str = "",
    employee_no: str = "",
    hire_date: str = "",
) -> str:
    """
    Search a person in 人力资源 and edit 工号 / 入职时间 in one save.

    Preconditions:
    - The browser should already be logged in to eTeams.
    - The function will enter/verify Org. Structure > 组织维护 > 人力资源.

    Args:
        driver: Existing Selenium Chrome driver. If omitted, uses this module's
            shared driver.
        person_name: Person name or unique keyword to search, for example
            ``xuyingtest人员05031651297305``.
        employee_no: New 工号. Leave empty to auto-generate a test value.
        hire_date: New 入职时间. Accepts ``yyyyMMdd``, ``yyyy-MM-dd`` or Chinese
            date text such as ``2016年3月27号``.
    """
    if not (person_name or "").strip():
        return "❌ 未提供要编辑的人员姓名/搜索关键词。"
    driver = driver or _get_driver()
    return _standalone_edit_person_employee_no_and_hire_date(
        driver,
        keyword=person_name,
        employee_no=employee_no,
        hire_date=hire_date,
    )

@function_tool(name_override="select_org_structure")
def select_org_structure_tool() -> str:
    """
    登录成功后，从右上角真实 eTeams 下拉菜单查找并选择 Org. Structure。

    不使用固定 URL 兜底，也不调整窗口宽度。
    """
    return select_org_structure()


@function_tool(name_override="create_public_group")
def create_public_group_tool(group_name: str = "") -> str:
    """
    在 Org. Structure 页面进入左侧「群组管理」，新建一个公共组。

    Args:
        group_name: 要创建的群组名称；留空时自动生成随机名称。
    """
    return create_public_group(group_name=group_name)


@function_tool(name_override="create_new_person")
def create_new_person_tool() -> str:
    """
    在已登录的 eTeams 中进入 Org. Structure，选择左侧「组织维护」、
    「人力资源」标签页，点击「新建人员」，填写必填字段和部门，
    保存前停留 5 秒，然后保存并搜索验证。
    """
    return create_new_person()


@function_tool(name_override="create_new_department")
def create_new_department_tool(department_name: str = "") -> str:
    """
    在已登录并进入 Org. Structure 的 eTeams 中，进入组织维护/分部信息，
    点击「新建部门」，填写部门名称并保存。

    Args:
        department_name: 要创建的部门名称；留空时自动生成 xuyingtest + 时间戳。
    """
    return create_new_department(department_name=department_name)


@function_tool(name_override="edit_person")
def edit_person_tool(
    person_name: str = "",
    employee_no: str = "",
    hire_date: str = "",
) -> str:
    """
    在已登录的 eTeams 中进入 Org. Structure > 组织维护 > 人力资源，
    搜索指定人员并同时编辑工号和入职时间，最后保存。

    Args:
        person_name: 要搜索/编辑的人员姓名或唯一关键词，例如
            xuyingtest人员05031651297305。
        employee_no: 新工号，例如 12345678。
        hire_date: 新入职时间，支持 20160327、2016-03-27、2016年3月27号。
    """
    return edit_person(
        person_name=person_name,
        employee_no=employee_no,
        hire_date=hire_date,
    )

@function_tool(name_override="login_and_create_new_person")
def login_and_create_new_person(
    username: str = DEFAULT_ACCOUNT,
    password: str = "",
) -> str:
    """
    从 eTeams Passport 执行完整流程：登录 -> 进入 Org. Structure ->
    组织维护 -> 人力资源 -> 新建人员 -> 填写必填字段和部门 ->
    保存前停留 5 秒 -> 保存 -> 搜索验证。

    如果 password 为空，会使用环境变量 ETEAMS_PASSWORD 或本地默认密码。

    Args:
        username: 登录账号/手机号。默认使用本地配置的 eTeams 账号。
        password: 登录密码。可留空以使用本地配置的密码。
    """
    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)
        privacy_result = _accept_privacy_terms_if_needed(driver)
        submit_result = _click_login_button(driver)

        outcome = _wait_for_login_result(driver)
        password_prompt_result = _decline_save_password_prompt_if_possible(driver)
        org_structure_result = ""
        person_result = ""

        if outcome.get("status") == "success":
            org_structure_result = select_org_structure(driver)
            if org_structure_result.startswith(("已选择", "已在", "已找到")):
                person_result = create_new_person(driver)

        current_url = driver.current_url
        current_title = driver.title
        messages = outcome.get("messages") or []

        if outcome.get("status") == "success":
            org_structure_success = org_structure_result.startswith(
                ("已选择", "已在", "已找到")
            )
            person_success = person_result.startswith("已完成新建人员流程")

            if org_structure_success and person_success:
                result_status = "✅ 登录成功，已进入 Org. Structure，并已完成新建人员。"
            elif org_structure_success:
                result_status = "⚠️ 登录成功，并已进入 Org. Structure，但未确认新建人员完成。"
            else:
                result_status = "⚠️ 登录成功，但未确认已进入/选择 Org. Structure。"

            result = (
                f"{result_status}"
                f" 已看到左上角登录用户，登录用户区域: {outcome.get('top_left_user')}。"
                f" {submit_result}"
                f" {org_structure_result}"
                f"{' ' + person_result if person_result else ''}"
                f" {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        elif outcome.get("status") == "failed" and messages:
            result = (
                "❌ 登录未完成。页面提示: "
                + "；".join(messages)
                + f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
            )
        elif outcome.get("page_changed"):
            result = (
                "⚠️ 页面已跳转或进入下一步，但未确认左上角登录用户，"
                "因此未按成功判定，也未执行新建人员。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        else:
            result = (
                "⚠️ 已提交登录，但结果不明确；未看到左上角登录用户，"
                "因此未执行新建人员。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )

        time.sleep(DEFAULT_CLOSE_DELAY_SECONDS)
        _close_driver()
        return f"{result}\n已停留 {DEFAULT_CLOSE_DELAY_SECONDS} 秒并关闭浏览器。"

    except TimeoutException:
        return "❌ 超时：无法找到登录表单元素或新建人员页面元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 登录并新建人员过程发生异常: {str(e)}"


@function_tool(name_override="login_and_create_new_department")
def login_and_create_new_department(
    department_name: str = "",
    username: str = DEFAULT_ACCOUNT,
    password: str = "",
) -> str:
    """
    完整新建部门流程：登录 -> 进入 Org. Structure -> 组织维护 ->
    分部信息 -> 新建部门 -> 填写部门名称 -> 保存 -> 停留 5 秒 -> 关闭浏览器。

    如果 password 为空，会使用环境变量 ETEAMS_PASSWORD 或本地默认密码。

    Args:
        department_name: 要创建的部门名称；留空时自动生成 xuyingtest + 时间戳。
        username: 登录账号/手机号。默认使用本地配置的 eTeams 账号。
        password: 登录密码。可留空以使用本地配置的密码。
    """
    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)
        privacy_result = _accept_privacy_terms_if_needed(driver)
        submit_result = _click_login_button(driver)

        outcome = _wait_for_login_result(driver)
        password_prompt_result = _decline_save_password_prompt_if_possible(driver)
        org_structure_result = ""
        department_result = ""

        if outcome.get("status") == "success":
            org_structure_result = select_org_structure(driver)
            if org_structure_result.startswith(("已选择", "已在", "已找到")):
                department_result = create_new_department(
                    driver,
                    department_name=department_name,
                )

        current_url = driver.current_url
        current_title = driver.title
        messages = outcome.get("messages") or []

        if outcome.get("status") == "success":
            org_structure_success = org_structure_result.startswith(
                ("已选择", "已在", "已找到")
            )
            department_success = (
                "✅ 部门创建成功" in department_result
                and "❌" not in department_result
                and "错误" not in department_result
            )

            if org_structure_success and department_success:
                result_status = "✅ 登录成功，已进入 Org. Structure，并已完成新建部门。"
            elif org_structure_success:
                result_status = "⚠️ 登录成功，并已进入 Org. Structure，但未确认新建部门完成。"
            else:
                result_status = "⚠️ 登录成功，但未确认已进入/选择 Org. Structure。"

            result = (
                f"{result_status}"
                f" 已看到左上角登录用户，登录用户区域: {outcome.get('top_left_user')}。"
                f" {submit_result}"
                f" {org_structure_result}"
                f"{' ' + department_result if department_result else ''}"
                f" {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        elif outcome.get("status") == "failed" and messages:
            result = (
                "❌ 登录未完成。页面提示: "
                + "；".join(messages)
                + f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
            )
        elif outcome.get("page_changed"):
            result = (
                "⚠️ 页面已跳转或进入下一步，但未确认左上角登录用户，"
                "因此未按成功判定，也未执行新建部门。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        else:
            result = (
                "⚠️ 已提交登录，但结果不明确；未看到左上角登录用户，"
                "因此未执行新建部门。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )

        time.sleep(DEFAULT_CLOSE_DELAY_SECONDS)
        _close_driver()
        return f"{result}\n已停留 {DEFAULT_CLOSE_DELAY_SECONDS} 秒并关闭浏览器。"

    except TimeoutException:
        return "❌ 超时：无法找到登录表单元素或新建部门页面元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 登录并新建部门过程发生异常: {str(e)}"


@function_tool(name_override="login_and_edit_person")
def login_and_edit_person(
    person_name: str = "",
    employee_no: str = "",
    hire_date: str = "",
    username: str = DEFAULT_ACCOUNT,
    password: str = "",
) -> str:
    """
    完整编辑人员流程：登录 -> 进入 Org. Structure -> 组织维护 ->
    人力资源 -> 搜索指定人员 -> 编辑工号和入职时间 -> 保存 ->
    停留 15 秒 -> 关闭浏览器。

    如果 password 为空，会使用环境变量 ETEAMS_PASSWORD 或本地默认密码。

    Args:
        person_name: 要搜索/编辑的人员姓名或唯一关键词，例如
            xuyingtest人员05031651297305。
        employee_no: 新工号，例如 12345678。
        hire_date: 新入职时间，支持 20160327、2016-03-27、2016年3月27号。
        username: 登录账号/手机号。默认使用本地配置的 eTeams 账号。
        password: 登录密码。可留空以使用本地配置的密码。
    """
    if not (person_name or "").strip():
        return "❌ 未提供要编辑的人员姓名/搜索关键词。"

    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)
        privacy_result = _accept_privacy_terms_if_needed(driver)
        submit_result = _click_login_button(driver)

        outcome = _wait_for_login_result(driver)
        password_prompt_result = _decline_save_password_prompt_if_possible(driver)
        edit_result = ""
        current_url = driver.current_url
        current_title = driver.title
        messages = outcome.get("messages") or []

        if outcome.get("status") == "success":
            edit_result = edit_person(
                driver,
                person_name=person_name,
                employee_no=employee_no,
                hire_date=hire_date,
            )
            current_url = driver.current_url
            current_title = driver.title

        if outcome.get("status") == "success":
            edit_success = edit_result.startswith("已完成")
            if edit_success:
                result_status = "✅ 登录成功，已完成人员编辑并保存。"
            else:
                result_status = "⚠️ 登录成功，但未确认人员编辑保存成功。"
            result = (
                f"{result_status}"
                f" 已看到左上角登录用户，登录用户区域: {outcome.get('top_left_user')}。"
                f" {submit_result}"
                f" {edit_result}"
                f" {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        elif outcome.get("status") == "failed" and messages:
            result = (
                "❌ 登录未完成。页面提示: "
                + "；".join(messages)
                + f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
            )
        elif outcome.get("page_changed"):
            result = (
                "⚠️ 页面已跳转或进入下一步，但未确认左上角登录用户，"
                "因此未执行人员编辑。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        else:
            result = (
                "⚠️ 已提交登录，但结果不明确；未看到左上角登录用户，"
                "因此未执行人员编辑。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )

        time.sleep(DEFAULT_CLOSE_DELAY_SECONDS)
        _close_driver()
        return f"{result}\n已停留 {DEFAULT_CLOSE_DELAY_SECONDS} 秒并关闭浏览器。"

    except TimeoutException:
        return "❌ 超时：无法找到登录表单元素或人员编辑页面元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 登录并编辑人员过程发生异常: {str(e)}"


def _open_eteams_login_page() -> str:
    driver = _get_driver()
    driver.get(TARGET_URL)
    _bring_driver_window_to_front(driver)
    language_result = _set_language_to_simplified_chinese(driver)
    _bring_driver_window_to_front(driver)
    return (
        f"浏览器已打开 eTeams Passport。{language_result}"
        f" 当前 URL: {driver.current_url}，页面标题: {driver.title}"
    )


@function_tool
def open_browser() -> str:
    """打开 Chrome 浏览器，进入 eTeams Passport，并将语言设置为简体中文。"""
    logger.info("调用 open_browser 工具")
    result = _open_eteams_login_page()
    logger.info(f"open_browser 结果: {result[:200]}")
    return result


@function_tool
def set_language_to_simplified_chinese() -> str:
    """将当前 eTeams Passport 登录页语言设置为简体中文。"""
    driver = _get_driver()
    return _set_language_to_simplified_chinese(driver)


@function_tool
def fill_login_form(username: str = DEFAULT_ACCOUNT, password: str = "") -> str:
    """
    在 eTeams Passport 登录页填写账号和密码，但不提交登录。

    如果 password 为空，会使用环境变量 ETEAMS_PASSWORD 或本地默认密码。

    Args:
        username: 登录账号/手机号。默认使用本地配置的 eTeams 账号。
        password: 登录密码。可留空以使用本地配置的密码。
    """
    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)

        return (
            f"✅ {language_result} 已填写账号 {_mask_account(username)}"
            " 和密码（已隐藏），尚未提交登录。"
        )
    except TimeoutException:
        return "❌ 超时：无法找到 eTeams Passport 登录表单元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 填写登录表单时发生异常: {str(e)}"


def _run_basic_login_test(username: str = DEFAULT_ACCOUNT, password: str = "") -> str:
    """Run only the basic Passport login test, without business operations."""
    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)
        privacy_result = _accept_privacy_terms_if_needed(driver)
        submit_result = _click_login_button(driver)

        outcome = _wait_for_login_result(driver)
        password_prompt_result = _decline_save_password_prompt_if_possible(driver)
        current_url = driver.current_url
        current_title = driver.title
        messages = outcome.get("messages") or []
        screenshot_result = ""
        try:
            screenshot_label = (
                "basic_login_success"
                if outcome.get("status") == "success"
                else "basic_login_result"
            )
            screenshot_result = _save_screenshot(driver, screenshot_label)
        except Exception as screenshot_error:
            screenshot_result = f"📸 截图失败：{screenshot_error}"

        if outcome.get("status") == "success":
            result_status = "✅ 基础登录测试通过：已看到左上角登录用户。"
            result = (
                f"{result_status}"
                f" 登录用户区域: {outcome.get('top_left_user')}。"
                f" {submit_result}"
                f" {language_result} {privacy_result} {password_prompt_result}"
                f" {screenshot_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        elif outcome.get("status") == "failed" and messages:
            result = (
                "❌ 登录未完成。页面提示: "
                + "；".join(messages)
                + f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                + f" {screenshot_result}"
            )
        elif outcome.get("page_changed"):
            result = (
                "⚠️ 页面已跳转或进入下一步，但未确认左上角登录用户，"
                "因此未按成功判定。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" {screenshot_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        else:
            result = (
                "⚠️ 已提交登录，但结果不明确；未看到左上角登录用户。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" {screenshot_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )

        time.sleep(DEFAULT_BASIC_LOGIN_CLOSE_DELAY_SECONDS)
        _close_driver()
        return f"{result}\n已停留 {DEFAULT_BASIC_LOGIN_CLOSE_DELAY_SECONDS} 秒并关闭浏览器。"

    except TimeoutException:
        return "❌ 超时：无法找到登录表单元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 登录过程发生异常: {str(e)}"


@function_tool
def login(username: str = DEFAULT_ACCOUNT, password: str = "") -> str:
    """
    基础登录测试：只验证 eTeams Passport 登录本身。

    流程：打开/确认 Passport -> 切换简体中文 -> 填写账号密码 ->
    勾选隐私协议 -> 点击登录 -> 看到左上角登录用户即判定成功 ->
    截图留证 -> 停留 5 秒 -> 关闭浏览器。

    不会进入 Org. Structure，不会创建公共组，也不会新建人员。

    如果 password 为空，会使用环境变量 ETEAMS_PASSWORD 或本地默认密码。

    Args:
        username: 登录账号/手机号。默认使用本地配置的 eTeams 账号。
        password: 登录密码。可留空以使用本地配置的密码。
    """
    logger.info(f"调用 login 工具，账号: {_mask_account(username)}")
    result = _run_basic_login_test(username=username, password=password)
    logger.info(f"login 结果: {result[:200]}...")
    return result


@function_tool(name_override="login_and_create_public_group")
def login_and_create_public_group(
    username: str = DEFAULT_ACCOUNT,
    password: str = "",
) -> str:
    """
    完整公共组流程：登录 -> 进入 Org. Structure -> 创建公共组 ->
    停留 5 秒 -> 关闭浏览器。

    仅当用户明确要求创建公共组时使用；基础登录测试请使用 login。
    """
    driver = _get_driver()
    try:
        if not _login_form_is_present(driver):
            driver.get(TARGET_URL)
            _bring_driver_window_to_front(driver)
        language_result = _set_language_to_simplified_chinese(driver)
        _bring_driver_window_to_front(driver)

        username = username or DEFAULT_ACCOUNT
        password_to_use = password or _DEFAULT_PASSWORD
        _fill_login_fields(driver, username, password_to_use)
        privacy_result = _accept_privacy_terms_if_needed(driver)
        submit_result = _click_login_button(driver)

        outcome = _wait_for_login_result(driver)
        password_prompt_result = _decline_save_password_prompt_if_possible(driver)
        org_structure_result = ""
        group_result = ""
        if outcome.get("status") == "success":
            org_structure_result = select_org_structure(driver)
            if org_structure_result.startswith(("已选择", "已在", "已找到")):
                group_result = create_public_group(driver)
        current_url = driver.current_url
        current_title = driver.title
        messages = outcome.get("messages") or []

        if outcome.get("status") == "success":
            org_structure_success = org_structure_result.startswith(
                ("已选择", "已在", "已找到")
            )
            group_success = group_result.startswith("已")
            if org_structure_success and group_success:
                result_status = "✅ 登录成功，已进入 Org. Structure，并已完成公共组创建/验证。"
            elif org_structure_success:
                result_status = "⚠️ 登录成功，并已进入 Org. Structure，但未确认公共组创建/验证。"
            else:
                result_status = "⚠️ 登录成功，但未确认已进入/选择 Org. Structure。"
            result = (
                f"{result_status}"
                f" 已看到左上角登录用户，登录用户区域: {outcome.get('top_left_user')}。"
                f" {submit_result}"
                f" {org_structure_result}"
                f"{' ' + group_result if group_result else ''}"
                f" {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        elif outcome.get("status") == "failed" and messages:
            result = (
                "❌ 登录未完成。页面提示: "
                + "；".join(messages)
                + f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
            )
        elif outcome.get("page_changed"):
            result = (
                "⚠️ 页面已跳转或进入下一步，但未确认左上角登录用户，"
                "因此未按成功判定，也未执行公共组创建。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )
        else:
            result = (
                "⚠️ 已提交登录，但结果不明确；未看到左上角登录用户，"
                "因此未执行公共组创建。"
                f" {submit_result} {language_result} {privacy_result} {password_prompt_result}"
                f" 当前 URL: {current_url}，页面标题: {current_title}"
            )

        time.sleep(DEFAULT_CLOSE_DELAY_SECONDS)
        _close_driver()
        return f"{result}\n已停留 {DEFAULT_CLOSE_DELAY_SECONDS} 秒并关闭浏览器。"

    except TimeoutException:
        return "❌ 超时：无法找到登录表单元素或公共组页面元素，请确认页面已正确加载。"
    except Exception as e:
        return f"❌ 登录并创建公共组过程发生异常: {str(e)}"


@function_tool
def get_current_page_info() -> str:
    """获取当前页面的 URL 和标题信息"""
    driver = _get_driver()
    return f"当前 URL: {driver.current_url}\n页面标题: {driver.title}"


@function_tool
def navigate_to_login_page() -> str:
    """导航回 eTeams Passport 登录页面，并将语言设置为简体中文。"""
    return _open_eteams_login_page()


@function_tool
def take_screenshot(label: str) -> str:
    """
    对当前浏览器页面截图并保存到项目目录下的 screenshots/ 文件夹。
    在每个测试用例执行完毕后调用，用于记录测试结果。

    Args:
        label: 截图标签，用于文件名，例如 "eteams_login_form" 或 "eteams_login_result"
    """
    driver = _get_driver()
    return _save_screenshot(driver, label)


@function_tool
def close_browser() -> str:
    """关闭浏览器并释放 WebDriver 资源"""
    if _close_driver():
        return "✅ 浏览器已关闭。"
    return "ℹ️ 浏览器本来就未打开。"
