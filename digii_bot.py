"""
Digii_Monday_Bot - Production-Ready DIGII ERP Administrative Automation Tool
Automates recurring Monday administrative tasks, report extractions, and compliance archiving.

Specialized Features & Zero-Write-Risk Demo Modules:
1. Real Chrome Persistent Context:
   - Launches system-installed Google Chrome (channel="chrome") with local persistent profile (./chrome_profile)
   - Removes automation-specific flags (navigator.webdriver masked, --enable-automation disabled) so anti-bot widgets recognize a standard browser
2. Smart Session Verification:
   - Navigates to https://poornima.digiicampus.com/V2/#/staffManagement
   - If already authenticated, proceeds directly to tasks and refreshes auth_state.json
   - If redirected to login, fills DIGII_USER & DIGII_PASS from .env and waits up to 45s (page.wait_for_url("**/V2/#/**", timeout=45000))
     for verification widgets to auto-resolve (or allow manual click if prompted)
3. Zero-Write-Risk Demo Tasks:
   - Task 1: REPORT_MODAL_DOWNLOAD (Faculty & Administrator Login / Not Logged In Reports)
   - Task 2: DIRECT_DOWNLOAD (Download Staff Data master export)
   - Task 3: FORM_DRY_RUN (Aggressively clears Chrome autofill -> populates demo record -> captures 04_Form_Dry_Run_Populated.png -> CLICKS CANCEL)
"""

import argparse
import datetime
import http.server
import logging
import os
import socketserver
import sys
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import (
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

# -----------------------------------------------------------------------------
# 1. Configuration & Path Management
# -----------------------------------------------------------------------------
class BotConfig:
    """Manages environment credentials and runtime configuration settings."""

    def __init__(self, env_path: Optional[str] = None):
        if env_path and os.path.exists(env_path):
            load_dotenv(dotenv_path=env_path)
        else:
            load_dotenv()

        # URLs
        self.login_url = os.getenv("LOGIN_URL", "https://poornima.digiicampus.com/V2/#/home").strip()
        self.dashboard_url = os.getenv("DASHBOARD_URL", "https://poornima.digiicampus.com/V2/#/staffManagement").strip()

        # Credentials
        self.user = os.getenv("DIGII_USER", "").strip()
        self.password = os.getenv("DIGII_PASS", "").strip()

        # Browser Settings
        headless_val = os.getenv("HEADLESS", "false").strip().lower()
        self.headless = headless_val in ("true", "1", "yes", "t")

        self.slow_mo_ms = int(os.getenv("SLOW_MO_MS", "300"))
        self.page_timeout_ms = int(os.getenv("PAGE_TIMEOUT_MS", "60000"))
        self.download_timeout_ms = int(os.getenv("DOWNLOAD_TIMEOUT_MS", "90000"))

        # Selectors (Optimized for DIGII ERP V2 / Poornima)
        self.selector_user = os.getenv(
            "SELECTOR_USERNAME",
            'input[name="registrationId"], input[name="username"], input[type="text"], #username',
        ).strip()
        self.selector_pass = os.getenv(
            "SELECTOR_PASSWORD",
            'input[name="password"], input[type="password"], #password',
        ).strip()
        self.selector_submit = os.getenv(
            "SELECTOR_SUBMIT",
            'button.signin-button, button:has-text("Sign in"), button:has-text("Login"), #btnLogin, button[type="submit"]',
        ).strip()
        self.selector_logged_in = os.getenv(
            "SELECTOR_LOGGED_IN_PROOF",
            'button:has-text("Reports"), button:has-text("Download Staff Data"), .user-profile, .dashboard, nav, [aria-label*="Logout"], button:has-text("Logout"), a[href*="staffManagement"]',
        ).strip()

    def is_placeholder_config(self) -> bool:
        """Determines if credentials or URLs are still set to placeholder values."""
        return (
            "digii.example.com" in self.login_url
            or not self.user
            or not self.password
            or self.user == "your_username_or_email"
            or self.password == "your_secret_password"
            or self.user == "admin@digii.demo"
        )


class RunArchiveManager:
    """Creates and coordinates timestamped run output directories."""

    def __init__(self, base_runs_dir: str = "Runs"):
        self.base_runs_dir = Path(base_runs_dir)
        self.timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
        self.run_dir = self.base_runs_dir / self.timestamp_str

        self.downloads_dir = self.run_dir / "raw_downloads"
        self.screenshots_dir = self.run_dir / "screenshots"
        self.report_path = self.run_dir / "Audit_Report.xlsx"
        self.log_path = self.run_dir / "run_execution.log"

        self.downloads_dir.mkdir(parents=True, exist_ok=True)
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)


# -----------------------------------------------------------------------------
# 2. Excel Configuration Reader
# -----------------------------------------------------------------------------
class ExcelConfigReader:
    """Loads and validates dynamic URL extraction tasks from Excel."""

    REQUIRED_COLUMNS = ["Task_Name", "Mode", "Target_URL", "Action_Target", "Active"]

    def __init__(self, config_path: str = "config/urls_config.xlsx"):
        self.config_path = Path(config_path)

    def load_active_tasks(self) -> List[Dict[str, Any]]:
        """Reads Excel file and returns all active tasks."""
        if not self.config_path.exists():
            raise FileNotFoundError(
                f"Configuration file not found: {self.config_path.resolve()}"
            )

        df = pd.read_excel(self.config_path, sheet_name=0)
        df.columns = [str(col).strip() for col in df.columns]

        missing = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            raise ValueError(
                f"Missing required columns in {self.config_path.name}: {missing}. "
                f"Expected: {self.REQUIRED_COLUMNS}"
            )

        active_tasks = []
        for idx, row in df.iterrows():
            raw_active = row.get("Active", False)
            is_active = False
            if pd.notna(raw_active):
                if isinstance(raw_active, (bool, int)):
                    is_active = bool(raw_active)
                elif isinstance(raw_active, str):
                    is_active = raw_active.strip().upper() in ("TRUE", "1", "YES", "Y", "T")

            if not is_active:
                continue

            task_name = str(row.get("Task_Name", f"Task_{idx+1}")).strip()
            mode = str(row.get("Mode", "DIRECT_DOWNLOAD")).strip().upper()
            target_url = str(row.get("Target_URL", "")).strip()
            action_target = str(row.get("Action_Target", "") if pd.notna(row.get("Action_Target")) else "").strip()

            active_tasks.append(
                {
                    "row_index": idx + 2,
                    "task_name": task_name,
                    "mode": mode,
                    "target_url": target_url,
                    "action_target": action_target,
                    "active": True,
                }
            )

        return active_tasks


# -----------------------------------------------------------------------------
# 3. Real Chrome Persistent Context & Hybrid Auth Manager
# -----------------------------------------------------------------------------
class AuthManager:
    """
    Manages Google Chrome persistent browser profile (./chrome_profile) with
    anti-bot mitigation flags and robust session restoration / verification.
    """

    def __init__(
        self,
        config: BotConfig,
        archive: RunArchiveManager,
        logger: logging.Logger,
        demo_port: Optional[int] = None,
    ):
        self.config = config
        self.archive = archive
        self.logger = logger
        self.demo_port = demo_port
        self.auth_state_path = Path("auth_state.json")
        self.profile_dir = Path("./chrome_profile").resolve()

    def create_persistent_browser_context(self, playwright: Playwright) -> BrowserContext:
        """
        Launches system-installed Google Chrome (channel="chrome") using a local persistent
        profile directory (./chrome_profile) and removes anti-bot automation flags.
        """
        self.profile_dir.mkdir(parents=True, exist_ok=True)

        # Anti-bot bypass flags
        launch_args = [
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-infobars",
            "--start-maximized",
        ]
        ignore_default_args = ["--enable-automation"]

        self.logger.info(
            f"[BROWSER] Launching Google Chrome (channel='chrome', profile='{self.profile_dir.name}', "
            f"headless={self.config.headless}, slow_mo={self.config.slow_mo_ms}ms)..."
        )

        try:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(self.profile_dir),
                channel="chrome",
                headless=self.config.headless,
                slow_mo=self.config.slow_mo_ms,
                viewport=None,
                accept_downloads=True,
                args=launch_args,
                ignore_default_args=ignore_default_args,
            )
        except Exception as ex:
            self.logger.warning(
                f"[BROWSER] Could not launch system Chrome channel ({ex}). Falling back to bundled Chromium engine..."
            )
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(self.profile_dir),
                headless=self.config.headless,
                slow_mo=self.config.slow_mo_ms,
                viewport=None,
                accept_downloads=True,
                args=launch_args,
                ignore_default_args=ignore_default_args,
            )

        # Mask navigator.webdriver in JavaScript environment
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        return context

    def ensure_authenticated_session(self, context: BrowserContext) -> BrowserContext:
        """
        Checks if already authenticated at https://poornima.digiicampus.com/V2/#/staffManagement.
        If already logged in, saves updated auth_state.json and proceeds.
        If redirected to login, fills .env credentials and waits up to 45s for verification.
        """
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(self.config.page_timeout_ms)

        check_url = self._resolve_target_url(self.config.dashboard_url)
        self.logger.info(f"[AUTH] Checking authentication status at target: {check_url}...")

        try:
            page.goto(check_url, wait_until="domcontentloaded", timeout=self.config.page_timeout_ms)
        except Exception as nav_err:
            self.logger.warning(f"[AUTH] Navigation warning ({nav_err}). Proceeding to examine page...")

        # Allow Angular SPA session initialization to settle
        page.wait_for_timeout(1000)

        # 1. Check if already authenticated
        if self._is_session_valid(page):
            self.logger.info(f"[AUTH] Already authenticated! Active session confirmed on {page.url}.")

            target_page = page
            # If DIGII landed or auto-redirected to #/feed, navigate to #/staffManagement via native UI clicks
            if "staffmanagement" not in page.url.lower():
                target_page = self._navigate_to_staff_management_via_ui(page)

            ss_path = self.archive.screenshots_dir / "auth_session_reused.png"
            target_page.screenshot(path=str(ss_path), full_page=True)

            # Save updated auth_state.json backup
            context.storage_state(path=str(self.auth_state_path))
            self.logger.info(f"[AUTH] Saved updated session state to {self.auth_state_path}")
            return context

        # 2. Redirected to login screen
        self.logger.info(f"[AUTH] Redirected to login screen ({page.url}). Entering credentials from .env...")

        # If on /home or not having fields yet, make sure login elements settle
        if not self._has_login_fields(page):
            login_url = self._resolve_target_url(self.config.login_url)
            if page.url.lower().rstrip("/") != login_url.lower().rstrip("/"):
                self.logger.info(f"[AUTH] Navigating to Login URL: {login_url}")
                try:
                    page.goto(login_url, wait_until="domcontentloaded", timeout=self.config.page_timeout_ms)
                except Exception:
                    pass
                page.wait_for_timeout(300)

        self._perform_login_and_await_verification(page, context)
        return context

    def _perform_login_and_await_verification(self, page: Page, context: BrowserContext):
        """Fills credentials and waits up to 45s for the verification widget to resolve."""
        # Fill Username (registrationId / username)
        user_selectors = [
            'input[name="registrationId"]',
            self.config.selector_user,
            'input[name="username"]',
            'input[type="text"]',
            '#username',
            'input[placeholder*="registration" i]',
            'input[placeholder*="user" i]',
            'input[placeholder*="email" i]',
        ]
        self._safe_clear_and_fill(page, user_selectors, self.config.user, "Username")

        # Fill Password
        pass_selectors = [
            'input[name="password"]',
            self.config.selector_pass,
            'input[type="password"]',
            '#password',
        ]
        self._safe_clear_and_fill(page, pass_selectors, self.config.password, "Password")

        # Click Submit
        submit_selectors = [
            'button.signin-button',
            'button:has-text("Sign in")',
            self.config.selector_submit,
            '#btnLogin',
            'button:has-text("Login")',
            'button[type="submit"]',
        ]
        submit_loc = self._find_visible_locator(page, ",".join(submit_selectors))
        if submit_loc:
            self.logger.info(f"[AUTH] Clicking login submit button [{submit_loc}]...")
            submit_loc.click()
        else:
            self.logger.warning("[AUTH] Submit button locator not found, pressing Enter on password...")
            page.keyboard.press("Enter")

        # Wait up to 45 seconds for login verification:
        # Pumps Playwright event loop every 300ms so frame navigation events are dispatched immediately.
        # Checks both URL departure from login/home and visible authenticated UI elements.
        self.logger.info(
            "[AUTH] Waiting up to 45s for login verification: checking URL transition or authenticated UI indicators..."
        )
        authenticated = False
        start_wait = time.time()

        auth_indicators = page.locator(
            'button:has-text("CONSOLE"), a:has-text("CONSOLE"), div:has-text("CONSOLE"), span:has-text("CONSOLE"), '
            '.user-profile, #userBadge, button:has-text("Reports")'
        )

        while time.time() - start_wait < 45:
            page.wait_for_timeout(300)
            cur = page.url.lower()

            # Check URL departure from login/home (e.g. redirected to #/feed or #/staffManagement)
            if ("/v2/#/" in cur or "/v2/" in cur) and not any(k in cur for k in ["/login", "#/login", "#/home", "/home"]):
                authenticated = True
                self.logger.info(f"[AUTH] Login confirmed via URL departure! URL: {page.url} in {time.time() - start_wait:.1f}s")
                break

            # Check presence of authenticated UI elements
            try:
                if auth_indicators.count() > 0 and auth_indicators.first.is_visible():
                    authenticated = True
                    self.logger.info(f"[AUTH] Login confirmed via visible UI element [{page.url}] in {time.time() - start_wait:.1f}s")
                    break
            except Exception:
                pass

        if not authenticated:
            if not self._is_session_valid(page):
                raise TimeoutError(
                    f"Login verification did not complete within 45s. Current URL: {page.url}"
                )

        # Post-Login Native UI Navigation: #/feed -> CONSOLE -> User Records -> Staff Management
        # Zero raw page.goto() or window.location.hash jumps!
        target_page = page
        if "staffmanagement" not in page.url.lower():
            target_page = self._navigate_to_staff_management_via_ui(page)

        self.logger.info(f"[AUTH] Verification succeeded! Authenticated on: {target_page.url}")

        # Capture proof
        ss_path = self.archive.screenshots_dir / "auth_login_success.png"
        target_page.screenshot(path=str(ss_path), full_page=True)
        self.logger.info(f"[PROOF] Auth login screenshot saved: {ss_path.name}")

        # Save updated auth_state.json
        context.storage_state(path=str(self.auth_state_path))
        self.logger.info(f"[AUTH] Saved updated session state to {self.auth_state_path}")

    def _navigate_to_staff_management_via_ui(self, page: Page) -> Page:
        """
        Navigates natively via DIGII ERP UI from #/feed to #/staffManagement:
        1. Clicks the top-header CONSOLE button.
        2. Waits for right-side drawer to open, clicks 'User Records'.
        3. Clicks 'Staff Management' box/card (which opens in a new tab) and waits for button:has-text("Reports"):visible.
        Leaves the original #/feed tab intact and returns the newly opened tab for task execution.
        Zero raw page.goto() or window.location.hash jumps.
        """
        if "staffmanagement" in page.url.lower():
            self.logger.info("[AUTH UI NAV] Already on #/staffManagement.")
            return page

        self.logger.info("[AUTH UI NAV] Navigating from #/feed to #/staffManagement via native UI clicks...")

        # 1. Wait for and click top-header CONSOLE button
        console_candidates = [
            'button:has-text("CONSOLE"):visible',
            'a:has-text("CONSOLE"):visible',
            'div:has-text("CONSOLE"):visible',
            'span:has-text("CONSOLE"):visible',
            '#btnConsole:visible',
        ]
        console_loc = None
        for sel in console_candidates:
            try:
                cand = page.locator(sel)
                if cand.count() > 0 and cand.last.is_visible():
                    console_loc = cand.last
                    break
            except Exception:
                continue

        if not console_loc:
            try:
                cand = page.get_by_text("CONSOLE", exact=True).last
                if cand.count() > 0 and cand.is_visible():
                    console_loc = cand
            except Exception:
                pass

        if not console_loc:
            console_loc = page.locator('button:has-text("CONSOLE"), a:has-text("CONSOLE"), div:has-text("CONSOLE"), span:has-text("CONSOLE")').last

        self.logger.info("[AUTH UI NAV] Waiting for and clicking CONSOLE button in top header...")
        try:
            console_loc.wait_for(state="visible", timeout=15000)
            console_loc.click()
        except Exception as c_err:
            self.logger.warning(f"[AUTH UI NAV] CONSOLE click error ({c_err}). Attempting fallback force click...")
            try:
                console_loc.click(force=True)
            except Exception:
                page.locator('button:has-text("CONSOLE"), a:has-text("CONSOLE")').last.click()

        page.wait_for_timeout(800)

        # 2. Wait for right-side slide-out drawer, locate 'User Records', scroll and click
        # If Staff Management is already visible, we can click it directly. Otherwise click User Records first.
        self.logger.info("[AUTH UI NAV] Locating 'User Records' or 'Staff Management' in drawer menu...")
        staff_mgmt_already = page.locator(':is(div, span, a, button, p, li):has-text("Staff Management"):visible')
        if staff_mgmt_already.count() == 0 or not staff_mgmt_already.last.is_visible():
            user_records_loc = None
            try:
                cand = page.locator(':is(div, span, a, button, p, li):has-text("User Records"):visible').last
                cand.wait_for(state="visible", timeout=10000)
                user_records_loc = cand
            except Exception:
                pass

            if not user_records_loc:
                candidates = [
                    '#drawerUserRecords:visible',
                    'div:has-text("User Records"):visible',
                    'span:has-text("User Records"):visible',
                    ':has-text("User Records"):visible',
                ]
                for sel in candidates:
                    try:
                        c = page.locator(sel)
                        if c.count() > 0 and c.last.is_visible():
                            user_records_loc = c.last
                            break
                    except Exception:
                        continue

            if user_records_loc:
                self.logger.info("[AUTH UI NAV] Clicking 'User Records'...")
                try:
                    user_records_loc.scroll_into_view_if_needed()
                except Exception:
                    pass
                try:
                    user_records_loc.click()
                except Exception:
                    user_records_loc.click(force=True)
                page.wait_for_timeout(800)
            else:
                self.logger.warning("[AUTH UI NAV] 'User Records' element not found in drawer.")

        # 3. Locate 'Staff Management' box/card, click, and wait for button:has-text("Reports"):visible
        self.logger.info("[AUTH UI NAV] Locating 'Staff Management' option/card...")
        staff_mgmt_loc = None
        try:
            cand = page.locator(':is(div, span, a, button, p, li):has-text("Staff Management"):visible').last
            cand.wait_for(state="visible", timeout=10000)
            staff_mgmt_loc = cand
        except Exception:
            pass

        if not staff_mgmt_loc:
            candidates = [
                '#drawerStaffMgmt:visible',
                'div:has-text("Staff Management"):visible',
                'span:has-text("Staff Management"):visible',
                ':has-text("Staff Management"):visible',
            ]
            for sel in candidates:
                try:
                    c = page.locator(sel)
                    if c.count() > 0 and c.last.is_visible():
                        staff_mgmt_loc = c.last
                        break
                except Exception:
                    continue

        target_page = page
        if staff_mgmt_loc:
            self.logger.info("[AUTH UI NAV] Clicking 'Staff Management' (expecting new tab to open)...")
            try:
                staff_mgmt_loc.scroll_into_view_if_needed()
            except Exception:
                pass

            new_page = None
            try:
                with page.context.expect_page(timeout=5000) as new_page_info:
                    try:
                        staff_mgmt_loc.click()
                    except Exception:
                        staff_mgmt_loc.click(force=True)
                new_page = new_page_info.value
                self.logger.info(f"[AUTH UI NAV] New tab opened successfully: {new_page.url}")
            except Exception:
                # If expect_page timed out, check if a new tab appeared in context.pages anyway
                page.wait_for_timeout(1000)
                if len(page.context.pages) > 1:
                    new_page = page.context.pages[-1]
                    self.logger.info(f"[AUTH UI NAV] New tab detected in context.pages: {new_page.url}")

            if new_page and not new_page.is_closed():
                target_page = new_page
            else:
                target_page = page
        else:
            self.logger.warning("[AUTH UI NAV] 'Staff Management' box/card not found.")
            if len(page.context.pages) > 1:
                target_page = page.context.pages[-1]

        target_page.bring_to_front()
        target_page.set_default_timeout(self.config.page_timeout_ms)

        # Wait for button:has-text("Reports"):visible on the new tab to confirm #/staffManagement has loaded
        self.logger.info("[AUTH UI NAV] Waiting for button:has-text('Reports'):visible on the target tab...")
        try:
            reports_btn = target_page.locator('button:has-text("Reports"), #btnReports').first
            reports_btn.wait_for(state="visible", timeout=15000)
            self.logger.info(f"[AUTH UI NAV] Verified #/staffManagement loaded successfully on target tab ({target_page.url})!")
            return target_page
        except Exception as ex:
            self.logger.warning(f"[AUTH UI NAV] Wait for Reports button timed out on target tab ({ex}). Current URL: {target_page.url}")
            return target_page

    def _is_session_valid(self, page: Page) -> bool:
        """Determines if the current page indicates an active authenticated session."""
        current_url = page.url.lower()

        # If on home login page or login hash
        if (
            "/login" in current_url
            or "#/login" in current_url
            or "#/home" in current_url
            or current_url.endswith("/v2/")
            or current_url.endswith("/v2/#/")
        ):
            try:
                if page.locator('input[name="registrationId"], button.signin-button').count() > 0:
                    return False
            except Exception:
                pass

        if "staffmanagement" in current_url:
            return True

        if "feed" in current_url:
            return True

        try:
            proof = page.locator(self.config.selector_logged_in)
            if proof.count() > 0 and proof.first.is_visible():
                return True
        except Exception:
            pass

        try:
            feed_proof = page.locator('.user-profile, nav, .nav-bar, #userBadge, button:has-text("CONSOLE"), text="CONSOLE"')
            if feed_proof.count() > 0 and feed_proof.first.is_visible():
                return True
        except Exception:
            pass

        return False

    def _has_login_fields(self, page: Page) -> bool:
        """Checks if login input fields are visible on current page."""
        try:
            user_input = page.locator('input[name="registrationId"], input[name="username"], input[type="text"], #username')
            pass_input = page.locator('input[name="password"], input[type="password"]')
            return (
                user_input.count() > 0
                and user_input.first.is_visible()
                and pass_input.count() > 0
                and pass_input.first.is_visible()
                and "staffmanagement/add" not in page.url.lower()
            )
        except Exception:
            return False

    def _safe_clear_and_fill(
        self,
        page: Page,
        selector_candidates: List[str],
        value: str,
        field_name: str,
    ):
        """Clicks, clears any pre-existing text or autofill, and fills the field."""
        for sel in selector_candidates:
            if not sel:
                continue
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    self.logger.info(f"[AUTH] Entering {field_name} into [{sel}]")
                    loc.first.click()
                    try:
                        loc.first.clear()
                    except Exception:
                        pass
                    loc.first.press("Control+A")
                    loc.first.press("Backspace")
                    time.sleep(0.05)
                    loc.first.fill(value)
                    return True
            except Exception:
                continue
        self.logger.warning(f"[AUTH] Warning: Field '{field_name}' not matched with candidates.")
        return False

    def _find_visible_locator(self, page: Page, comma_separated_selectors: str):
        """Finds the first visible Playwright locator from candidates."""
        candidates = [s.strip() for s in comma_separated_selectors.split(",") if s.strip()]
        for sel in candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    return loc.first
            except Exception:
                continue
        return None

    def _resolve_target_url(self, raw_url: str) -> str:
        """Translates target URL to active mock server during demo runs."""
        if self.demo_port:
            if "poornima.digiicampus.com" in raw_url or "digii.example.com" in raw_url:
                if "#" in raw_url:
                    fragment = raw_url.split("#", 1)[1]
                    return f"http://127.0.0.1:{self.demo_port}/V2/#{fragment}"
                return f"http://127.0.0.1:{self.demo_port}/V2/#/staffManagement"
        return raw_url


# -----------------------------------------------------------------------------
# 4. Built-in Local Demo Server (Poornima DIGII ERP Simulation)
# -----------------------------------------------------------------------------
DEMO_HTML_PORTAL = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Poornima DIGII ERP - Staff Management Demo</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; background-color: #f0f2f5; color: #2d3748; }
        .header { background: #003366; color: white; padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .header h1 { margin: 0; font-size: 19px; font-weight: 600; display: flex; align-items: center; gap: 10px; }
        .user-badge { background: #004d99; padding: 5px 12px; border-radius: 4px; font-size: 13px; }
        .nav-bar { background: #ffffff; border-bottom: 1px solid #e2e8f0; padding: 10px 28px; display: flex; gap: 15px; font-size: 14px; }
        .nav-link { color: #003366; text-decoration: none; font-weight: 500; cursor: pointer; }
        .container { padding: 25px 28px; max-width: 1250px; margin: 0 auto; }
        .card { background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 24px; margin-bottom: 20px; }
        .toolbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }
        .btn-group { display: flex; gap: 10px; }
        .btn { padding: 9px 16px; border-radius: 4px; font-weight: 600; font-size: 13px; cursor: pointer; border: none; display: inline-flex; align-items: center; gap: 6px; }
        .btn-primary { background: #003366; color: white; }
        .btn-primary:hover { background: #002244; }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #218838; }
        .btn-secondary { background: #6c757d; color: white; }
        .btn-secondary:hover { background: #5a6268; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
        th, td { text-align: left; padding: 11px 12px; border-bottom: 1px solid #edf2f7; }
        th { background: #f8fafc; font-weight: 600; color: #4a5568; }
        .badge { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
        .badge-active { background: #d4edda; color: #155724; }
        
        .modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: none; justify-content: center; align-items: center; z-index: 1000; }
        .modal { background: white; border-radius: 8px; width: 500px; max-width: 90%; box-shadow: 0 10px 25px rgba(0,0,0,0.2); overflow: hidden; }
        .modal-header { background: #003366; color: white; padding: 15px 20px; display: flex; justify-content: space-between; align-items: center; }
        .modal-header h3 { margin: 0; font-size: 16px; }
        .modal-body { padding: 20px; }
        .report-item { padding: 12px 15px; border: 1px solid #e2e8f0; border-radius: 6px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; justify-content: space-between; font-weight: 500; }
        .report-item:hover { background: #ebf8ff; border-color: #3182ce; color: #2b6cb0; }
        
        .form-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-top: 15px; }
        .form-group { display: flex; flex-direction: column; gap: 5px; }
        .form-group label { font-size: 12px; font-weight: 600; color: #4a5568; }
        .form-group input, .form-group select { padding: 9px 12px; border: 1px solid #cbd5e0; border-radius: 4px; font-size: 13px; }
        .radio-group { display: flex; gap: 20px; align-items: center; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="header">
        <h1><span>🏛️</span> Poornima University &bull; DIGII ERP Portal</h1>
        <div style="display:flex; align-items:center; gap:12px;">
            <button id="btnConsole" class="btn" style="background:#dc3545; color:white; font-weight:bold; padding:6px 14px; border-radius:4px;" onclick="document.getElementById('consoleDrawer').style.display='block';">CONSOLE</button>
            <div class="user-badge" id="userBadge">&#128100; Administrator: demo_admin@poornima.org</div>
        </div>
    </div>
    <div id="consoleDrawer" style="display:none; position:fixed; top:0; right:0; width:300px; height:100%; background:white; box-shadow:-4px 0 12px rgba(0,0,0,0.15); z-index:2000; padding:20px; box-sizing:border-box;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #eee; padding-bottom:10px; margin-bottom:15px;">
            <h3 style="margin:0; font-size:16px; color:#003366;">Console Menu</h3>
            <button style="border:none; background:none; font-size:18px; cursor:pointer;" onclick="document.getElementById('consoleDrawer').style.display='none';">&times;</button>
        </div>
        <div style="margin-bottom:12px;">
            <div id="drawerUserRecords" style="font-weight:600; padding:10px; cursor:pointer; background:#f0f4f8; border-radius:4px; margin-bottom:8px;" onclick="document.getElementById('drawerStaffMgmt').style.display='block';">User Records</div>
            <div id="drawerStaffMgmt" style="display:block; padding:10px 15px; cursor:pointer; color:#003366; background:#e2e8f0; border-radius:4px;" onclick="document.getElementById('consoleDrawer').style.display='none'; navigate('#/staffManagement');">Staff Management</div>
        </div>
    </div>
    <div class="nav-bar">
        <a class="nav-link" onclick="navigate('#/staffManagement')">👥 Staff Management</a>
        <a class="nav-link" onclick="navigate('#/staffManagement/add')">➕ Add New Staff</a>
    </div>

    <div class="container" id="mainContainer"></div>

    <div class="modal-overlay" id="reportModal">
        <div class="modal">
            <div class="modal-header">
                <h3 id="reportModalTitle">Select a Report to Download</h3>
                <button style="background:none; border:none; color:white; font-size:18px; cursor:pointer;" onclick="closeReportModal()">&times;</button>
            </div>
            <div class="modal-body" id="reportModalBody">
                <div class="report-item" id="reportItemLogin" onclick="showReportFilterView('Faculty & Administrator Login Report')">
                    Faculty & Administrator Login Report
                </div>
                <div class="report-item" id="reportItemNotLogin" onclick="showReportFilterView('Faculty & Administrator Not Logged In Report')">
                    Faculty & Administrator Not Logged In Report
                </div>
                <div class="report-item" id="reportItemLicense" onclick="showReportFilterView('Faculty & Administrator License Report')">
                    Faculty & Administrator License Report
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button class="btn btn-secondary" onclick="closeReportModal()">Close</button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="staffDataModal">
        <div class="modal">
            <div class="modal-header">
                <h3>Download Staff Data</h3>
                <button style="background:none; border:none; color:white; font-size:18px; cursor:pointer;" onclick="closeStaffDataModal()">&times;</button>
            </div>
            <div class="modal-body">
                <p style="font-size:13px; color:#4a5568;">Select fields to include in export:</p>
                <div style="display:flex; flex-direction:column; gap:8px; margin:15px 0;">
                    <label><input type="checkbox" id="selectAllCheckbox" checked /> <strong>Select All</strong></label>
                    <label style="margin-left:15px;"><input type="checkbox" id="personalDetailsCheckbox" checked /> Personal Details</label>
                    <label style="margin-left:15px;"><input type="checkbox" checked /> Academic Details</label>
                </div>
                <div style="text-align:right; margin-top:20px; display:flex; justify-content:flex-end; gap:10px;">
                    <button type="button" class="btn btn-secondary" onclick="closeStaffDataModal()">Cancel</button>
                    <button type="button" class="btn btn-danger" style="background:#dc3545; color:white;" onclick="triggerStaffEmailDispatch()">Download</button>
                </div>
            </div>
        </div>
    </div>
    <div id="emailToast" class="toast-success" style="display:none; position:fixed; bottom:20px; right:20px; background:#28a745; color:white; padding:12px 20px; border-radius:6px; box-shadow:0 4px 12px rgba(0,0,0,0.15); z-index:9999; font-size:13px;">
        You will receive this report shortly to your email address.
    </div>

    <script>
        function navigate(hash) {
            window.location.hash = hash;
            renderPage();
        }

        function renderPage() {
            const hash = window.location.hash || '';
            const path = window.location.pathname;
            const container = document.getElementById('mainContainer');

            if (hash.includes('/login') || path.includes('/login') || hash.includes('/home')) {
                container.innerHTML = `
                    <div style="max-width:400px; margin:60px auto; background:white; padding:35px; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1);">
                        <h2 style="color:#003366; margin-top:0;">Poornima DIGII Sign-In</h2>
                        <p style="color:#666; font-size:13px;">Enterprise Institutional Authentication</p>
                        <form onsubmit="event.preventDefault(); window.location.hash = '#/feed'; renderPage();">
                            <div style="margin-bottom:15px;">
                                <label style="display:block; font-size:12px; font-weight:600; margin-bottom:5px;">Username / Registration ID</label>
                                <input type="text" id="username" name="registrationId" style="width:100%; padding:9px; border:1px solid #ccc; border-radius:4px; box-sizing:border-box;" value="demo_admin" required />
                            </div>
                            <div style="margin-bottom:20px;">
                                <label style="display:block; font-size:12px; font-weight:600; margin-bottom:5px;">Password</label>
                                <input type="password" id="password" name="password" style="width:100%; padding:9px; border:1px solid #ccc; border-radius:4px; box-sizing:border-box;" value="demo_password" required />
                            </div>
                            <button type="submit" id="btnLogin" class="btn btn-primary signin-button" style="width:100%; padding:10px; justify-content:center;">Sign In to DIGII</button>
                        </form>
                    </div>
                `;
                return;
            }

            if (hash.includes('/feed')) {
                container.innerHTML = `
                    <div class="card">
                        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #edf2f7; padding-bottom:12px; margin-bottom:15px;">
                            <h2 style="margin:0; font-size:18px; color:#003366;">Community Feed</h2>
                            <span class="badge" style="background:#e2e8f0; color:#4a5568;">Social Feed</span>
                        </div>
                        <p style="color:#555; font-size:14px;">Welcome to DIGII Campus Feed. Click <strong>CONSOLE</strong> in the top header to access administrative consoles.</p>
                    </div>
                `;
                return;
            }

            if (hash.includes('staffManagement/add')) {
                container.innerHTML = `
                    <div class="card">
                        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #edf2f7; padding-bottom:12px; margin-bottom:15px;">
                            <h2 style="margin:0; font-size:18px; color:#003366;">Add New Staff / Faculty Record</h2>
                            <span class="badge" style="background:#e2e8f0; color:#4a5568;">Zero-Write Dry Run Mode</span>
                        </div>
                        <form id="addStaffForm" onsubmit="event.preventDefault();">
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Employee ID *</label>
                                    <input type="text" id="employeeId" name="employeeId" placeholder="Employee ID" value="AUTO_SAVED_EMP_888" required />
                                </div>
                                <div class="form-group">
                                    <label>User Type *</label>
                                    <select id="userType" name="userType" required>
                                        <option value="">Select User Type</option>
                                        <option value="Faculty">Faculty</option>
                                        <option value="Administrator">Administrator</option>
                                        <option value="Staff">Staff</option>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Department *</label>
                                    <select id="department" name="department" required>
                                        <option value="">Select Department</option>
                                        <option value="CSE">Computer Science & Engineering</option>
                                        <option value="ECE">Electronics & Communication</option>
                                        <option value="ME">Mechanical Engineering</option>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label>Gender *</label>
                                    <div class="radio-group">
                                        <label><input type="radio" name="gender" value="Male" id="genderMale" /> Male</label>
                                        <label><input type="radio" name="gender" value="Female" id="genderFemale" /> Female</label>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label>First Name *</label>
                                    <input type="text" id="firstName" name="firstName" placeholder="First Name" required />
                                </div>
                                <div class="form-group">
                                    <label>Last name *</label>
                                    <input type="text" id="lastName" name="lastName" placeholder="Last name" required />
                                </div>
                                <div class="form-group">
                                    <label>Email ID *</label>
                                    <input type="email" id="email" name="email" placeholder="Email ID" required />
                                </div>
                                <div class="form-group">
                                    <label>Phone Number *</label>
                                    <input type="tel" id="phoneNumber" name="phoneNumber" placeholder="Phone Number" required />
                                </div>
                                <div class="form-group">
                                    <label>Password *</label>
                                    <input type="password" id="password" name="password" placeholder="Password" value="PreExistingAutofill123" required />
                                </div>
                            </div>
                            <div style="margin-top:25px; padding-top:15px; border-top:1px solid #edf2f7; display:flex; justify-content:flex-end; gap:12px;">
                                <button type="button" class="btn btn-secondary" id="btnCancel" onclick="handleCancel()">Cancel</button>
                                <button type="button" class="btn btn-primary" id="btnCreateUser">Create User</button>
                            </div>
                        </form>
                    </div>
                `;
            } else {
                container.innerHTML = `
                    <div class="card">
                        <div class="toolbar">
                            <div>
                                <h2 style="margin:0; font-size:18px; color:#003366;">Staff & Faculty Management</h2>
                                <p style="margin:4px 0 0 0; color:#718096; font-size:13px;">Manage academic rosters, reports, and administrative credentials</p>
                            </div>
                            <div class="btn-group">
                                <button class="btn btn-primary" id="btnReports" onclick="openReportModal()">
                                    <span>📑</span> Reports
                                </button>
                                <button class="btn btn-success" id="btnDownloadStaff" onclick="openStaffDataModal()">
                                    <span>📥</span> Download Staff Data
                                </button>
                                <div style="position:relative; display:inline-block;">
                                    <button class="btn btn-secondary" id="btnActions" onclick="const d=document.getElementById('actionsDropdown'); d.style.display=(d.style.display==='block'?'none':'block');">
                                        Actions &#9662;
                                    </button>
                                    <div id="actionsDropdown" style="display:none; position:absolute; top:100%; right:0; background:white; border:1px solid #cbd5e0; border-radius:4px; box-shadow:0 4px 12px rgba(0,0,0,0.15); min-width:140px; z-index:1000;">
                                        <a style="display:block; padding:8px 14px; color:#2d3748; text-decoration:none; cursor:pointer;" onclick="navigate('#/staffManagement/add'); document.getElementById('actionsDropdown').style.display='none';">Add Staff</a>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <table>
                            <thead>
                                <tr>
                                    <th>Employee ID</th><th>Staff Name</th><th>Department</th><th>Role</th><th>Email ID</th><th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr><td>#PU-FAC-101</td><td>Dr. Rajesh Sharma</td><td>Computer Science</td><td>Professor</td><td>rajesh.sharma@poornima.org</td><td><span class="badge badge-active">Active</span></td></tr>
                                <tr><td>#PU-FAC-102</td><td>Dr. Sunita Verma</td><td>Electronics & Comm</td><td>Associate Prof</td><td>sunita.verma@poornima.org</td><td><span class="badge badge-active">Active</span></td></tr>
                                <tr><td>#PU-ADM-044</td><td>Vikramaditya Rathore</td><td>Administration</td><td>Registrar</td><td>registrar@poornima.org</td><td><span class="badge badge-active">Active</span></td></tr>
                            </tbody>
                        </table>
                    </div>
                `;
            }
        }

        function openReportModal() {
            document.getElementById('reportModalTitle').innerText = 'Select a Report to Download';
            document.getElementById('reportModalBody').innerHTML = `
                <div class="report-item" id="reportItemLogin" onclick="showReportFilterView('Faculty & Administrator Login Report')">
                    Faculty & Administrator Login Report
                </div>
                <div class="report-item" id="reportItemNotLogin" onclick="showReportFilterView('Faculty & Administrator Not Logged In Report')">
                    Faculty & Administrator Not Logged In Report
                </div>
                <div class="report-item" id="reportItemLicense" onclick="showReportFilterView('Faculty & Administrator License Report')">
                    Faculty & Administrator License Report
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button class="btn btn-secondary" onclick="closeReportModal()">Close</button>
                </div>
            `;
            document.getElementById('reportModal').style.display = 'flex';
        }

        function showReportFilterView(reportName) {
            const cleanTitle = reportName.replace('.xlsx', '').replace(/_/g, ' ');
            document.getElementById('reportModalTitle').innerText = cleanTitle;
            document.getElementById('reportModalBody').innerHTML = `
                <p style="font-size:13px; color:#4a5568; margin-top:0;">Apply filters and select columns, then download</p>
                <div style="margin:12px 0; position:relative;">
                    <label style="display:block; font-size:12px; font-weight:600; margin-bottom:4px;">USER TYPE *</label>
                    <div id="userTypeTrigger" style="padding:7px 10px; width:100%; border:1px solid #cbd5e0; border-radius:4px; font-size:13px; cursor:pointer; background:white; box-sizing:border-box;" onclick="document.getElementById('userTypePopover').style.display = 'block';">
                        All User Type
                    </div>
                    <div id="userTypePopover" style="display:none; position:absolute; top:100%; left:0; width:100%; background:white; border:1px solid #cbd5e0; border-radius:4px; box-shadow:0 4px 12px rgba(0,0,0,0.15); padding:10px; z-index:1001; box-sizing:border-box;">
                        <input type="text" placeholder="Search..." style="width:100%; padding:5px 8px; margin-bottom:8px; border:1px solid #ccc; border-radius:3px; box-sizing:border-box;" />
                        <div style="display:flex; flex-direction:column; gap:6px; font-size:13px;">
                            <label><input type="checkbox" id="userTypeSelectAll" checked /> Select all</label>
                            <label><input type="checkbox" id="userTypeAdmin" checked /> administrator</label>
                            <label><input type="checkbox" id="userTypeFaculty" checked /> faculty</label>
                        </div>
                        <div style="text-align:right; margin-top:8px;">
                            <button type="button" class="btn btn-danger" style="background:#dc3545; color:white; padding:4px 12px; font-size:12px;" onclick="document.getElementById('userTypePopover').style.display = 'none';">OK</button>
                        </div>
                    </div>
                    <select id="userTypeSelect" name="userType" style="display:none;">
                        <option value="All" selected>All</option>
                    </select>
                </div>
                <div style="margin:15px 0;">
                    <label style="display:block; font-size:12px; font-weight:600; margin-bottom:6px;">Export Format</label>
                    <button type="button" class="btn btn-secondary" style="border:2px solid #003366; background:#e6f0fa; color:#003366; font-weight:bold;">XLSX</button>
                    <button type="button" class="btn btn-secondary" style="margin-left:8px;">CSV</button>
                </div>
                <div style="text-align:right; margin-top:20px; display:flex; justify-content:flex-end; gap:10px;">
                    <button type="button" class="btn btn-secondary" onclick="closeReportModal()">Cancel</button>
                    <button type="button" class="btn btn-danger" style="background:#dc3545; color:white;" onclick="showReportReadyModal('` + cleanTitle + `.xlsx')">Download XLSX</button>
                </div>
            `;
        }

        function showReportReadyModal(filename) {
            document.getElementById('reportModalTitle').innerText = 'Your Report is Ready';
            document.getElementById('reportModalBody').innerHTML = `
                <div style="text-align:center; padding:15px 0;">
                    <h3 style="color:#28a745; margin-bottom:10px;">Your report is ready!</h3>
                    <p style="color:#555; font-size:13px;">The requested report data has been compiled successfully.</p>
                    <div style="margin:20px 0;">
                        <a href="#" class="btn btn-primary" onclick="event.preventDefault(); triggerDownload('` + filename + `')">Download XLSX</a>
                    </div>
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button type="button" class="btn btn-secondary" onclick="closeReportModal()">Close</button>
                </div>
            `;
        }

        function closeReportModal() {
            document.getElementById('reportModal').style.display = 'none';
        }

        function openStaffDataModal() {
            document.getElementById('staffDataModal').style.display = 'flex';
        }

        function closeStaffDataModal() {
            document.getElementById('staffDataModal').style.display = 'none';
        }

        function triggerStaffEmailDispatch() {
            closeStaffDataModal();
            const toast = document.getElementById('emailToast');
            if (toast) {
                toast.style.display = 'block';
                setTimeout(() => { toast.style.display = 'none'; }, 6000);
            }
        }

        function handleCancel() {
            navigate('#/staffManagement');
        }

        function triggerDownload(filename) {
            const content = "Employee_ID,Name,Department,Role,Status\\n#PU-101,Dr. Rajesh Sharma,CSE,Faculty,Active\\n#PU-102,Dr. Sunita Verma,ECE,Faculty,Active";
            const blob = new Blob([content], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                try { document.body.removeChild(a); } catch(e){}
                window.URL.revokeObjectURL(url);
            }, 3000);
            closeReportModal();
        }

        window.addEventListener('hashchange', renderPage);
        window.addEventListener('DOMContentLoaded', renderPage);
    </script>
</body>
</html>
"""

class DemoServerHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(DEMO_HTML_PORTAL.encode("utf-8"))

    def log_message(self, format, *args):
        pass


class LocalDemoServer:
    """Spins up a lightweight local mock server on an available port for zero-config verification."""

    def __init__(self):
        self.port = 8765
        self.httpd = None
        self.thread = None

    def start(self) -> int:
        socketserver.TCPServer.allow_reuse_address = True
        for p in range(8765, 8790):
            try:
                self.httpd = socketserver.TCPServer(("127.0.0.1", p), DemoServerHandler)
                self.port = p
                break
            except OSError:
                continue

        if not self.httpd:
            raise RuntimeError("Could not find free port for Demo Server.")

        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()
        return self.port

    def stop(self):
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()


# -----------------------------------------------------------------------------
# 5. Extraction & Execution Engine (Zero-Risk & Resilient)
# -----------------------------------------------------------------------------
class DigiiAutomationEngine:
    """
    Orchestrates resilient navigation, SPA settlement, report extractions,
    and safe zero-write-risk dry-run form automation.
    """

    def __init__(
        self,
        config: BotConfig,
        archive: RunArchiveManager,
        logger: logging.Logger,
        demo_port: Optional[int] = None,
    ):
        self.config = config
        self.archive = archive
        self.logger = logger
        self.demo_port = demo_port
        self.audit_records: List[Dict[str, Any]] = []

    def _get_target_page(self, context: BrowserContext) -> Page:
        """
        Returns the active tab where tasks should run.
        Prioritizes the new tab (e.g. #/staffManagement) opened from the #/feed tab,
        leaving the original #/feed tab completely untouched.
        """
        # Look for any open page with staffmanagement in URL
        for p in reversed(context.pages):
            if not p.is_closed() and "staffmanagement" in p.url.lower():
                p.bring_to_front()
                return p

        # If multiple tabs exist, prefer the newest tab (last in context.pages)
        if len(context.pages) > 1:
            for p in reversed(context.pages):
                if not p.is_closed():
                    p.bring_to_front()
                    return p

        return context.pages[0] if context.pages else context.new_page()

    def execute_tasks(self, context: BrowserContext, tasks: List[Dict[str, Any]]):
        """Processes each task with isolated try/except fault tolerance."""
        page = self._get_target_page(context)
        page.set_default_timeout(self.config.page_timeout_ms)
        page.bring_to_front()

        total = len(tasks)
        self.logger.info(f"[ENGINE] Starting processing of {total} active task(s) on tab [{page.url}]...")

        for idx, task in enumerate(tasks, start=1):
            if page.is_closed():
                page = self._get_target_page(context)
                page.set_default_timeout(self.config.page_timeout_ms)
                page.bring_to_front()

            task_name = task["task_name"]
            mode = task["mode"]
            raw_url = task["target_url"]
            action_target = task["action_target"]

            target_url = self._resolve_target_url(raw_url)

            self.logger.info(f"\n--- [TASK {idx}/{total}] {task_name} | Mode: {mode} ---")
            self.logger.info(f"Target URL: {target_url}")

            timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            downloaded_file_rel: Optional[str] = None
            screenshot_proof_rel: Optional[str] = None
            status = "FAILED"
            error_notes = ""

            try:
                # 1. Resilient Navigation
                if mode == "FORM_DRY_RUN":
                    # Task 4: UI navigation (Actions -> Add Staff on #/staffManagement)
                    self.logger.info("[FORM DRY-RUN] Navigating to Add Staff via UI clicks (Actions -> Add Staff)...")
                    self._navigate_to_add_staff_via_ui(page)
                else:
                    if "staffmanagement" in page.url.lower() and "staffmanagement" in target_url.lower() and not page.url.lower().endswith("/add"):
                        self.logger.info(f"Already on target {target_url} on active tab; skipping redundant page.goto.")
                    else:
                        self.logger.info(f"Navigating to {target_url} (wait_until='domcontentloaded')...")
                        page.goto(target_url, wait_until="domcontentloaded", timeout=self.config.page_timeout_ms)

                # 2. Dynamic SPA settlement
                self._wait_for_spa_settlement(page)

                # 3. Action / Mode Execution
                if mode == "REPORT_MODAL_DOWNLOAD":
                    downloaded_file_rel = self._handle_report_modal_download(page, task_name, action_target)
                    ss_filename = f"{task_name}_success.png"
                    ss_full_path = self.archive.screenshots_dir / ss_filename
                    if not ss_full_path.exists():
                        page.screenshot(path=str(ss_full_path), full_page=True)
                    screenshot_proof_rel = f"screenshots/{ss_filename}"
                    status = "SUCCESS" if not self.demo_port else "DRY_RUN_SUCCESS"

                elif mode == "MODAL_EMAIL_DISPATCH":
                    downloaded_file_rel, screenshot_proof_rel = self._handle_modal_email_dispatch(page, task_name, action_target)
                    status = "SUCCESS"

                elif mode in ("DIRECT_DOWNLOAD", "EXPORT"):
                    downloaded_file_rel = self._handle_direct_download(page, task_name, action_target)
                    ss_filename = f"{task_name}_success.png"
                    ss_full_path = self.archive.screenshots_dir / ss_filename
                    page.screenshot(path=str(ss_full_path), full_page=True)
                    screenshot_proof_rel = f"screenshots/{ss_filename}"
                    status = "SUCCESS" if not self.demo_port else "DRY_RUN_SUCCESS"

                elif mode == "FORM_DRY_RUN":
                    screenshot_proof_rel = self._handle_form_dry_run(page, task_name, action_target)
                    status = "DRY_RUN_SUCCESS"
                    downloaded_file_rel = "N/A (Form Dry Run)"

                elif mode == "NAVIGATE_ONLY":
                    self.logger.info("NAVIGATE_ONLY mode: Page loaded and settled successfully.")
                    ss_filename = f"{task_name}_success.png"
                    ss_full_path = self.archive.screenshots_dir / ss_filename
                    page.screenshot(path=str(ss_full_path), full_page=True)
                    screenshot_proof_rel = f"screenshots/{ss_filename}"
                    status = "SUCCESS" if not self.demo_port else "DRY_RUN_SUCCESS"

                else:
                    self.logger.warning(f"Unrecognized mode '{mode}'. Falling back to direct export.")
                    downloaded_file_rel = self._handle_direct_download(page, task_name, action_target)
                    ss_filename = f"{task_name}_success.png"
                    ss_full_path = self.archive.screenshots_dir / ss_filename
                    page.screenshot(path=str(ss_full_path), full_page=True)
                    screenshot_proof_rel = f"screenshots/{ss_filename}"
                    status = "SUCCESS" if not self.demo_port else "DRY_RUN_SUCCESS"

                self.logger.info(f"[PROOF] Screenshot captured: {screenshot_proof_rel}")

            except Exception as ex:
                error_notes = f"{type(ex).__name__}: {str(ex).strip()}"
                self.logger.error(f"[ERROR] Task '{task_name}' encountered an exception: {error_notes}")

                try:
                    fail_ss_filename = f"{task_name}_failed.png"
                    fail_ss_path = self.archive.screenshots_dir / fail_ss_filename
                    page.screenshot(path=str(fail_ss_path), full_page=True)
                    screenshot_proof_rel = f"screenshots/{fail_ss_filename}"
                    self.logger.info(f"[PROOF] Failure screenshot captured: {screenshot_proof_rel}")
                except Exception as ss_err:
                    self.logger.warning(f"Could not capture failure screenshot: {ss_err}")

            finally:
                self.audit_records.append(
                    {
                        "Timestamp": timestamp_str,
                        "Task_Name": task_name,
                        "Target_URL": raw_url,
                        "Status": status,
                        "Downloaded_File": downloaded_file_rel or "N/A",
                        "Screenshot_Proof": screenshot_proof_rel or "N/A",
                        "Error_Notes": error_notes,
                    }
                )

    def _resolve_target_url(self, raw_url: str) -> str:
        """Translates live or template URLs to the active mock server during demo runs."""
        if self.demo_port:
            if "poornima.digiicampus.com" in raw_url or "digii.example.com" in raw_url:
                if "#" in raw_url:
                    fragment = raw_url.split("#", 1)[1]
                    return f"http://127.0.0.1:{self.demo_port}/V2/#{fragment}"
                return f"http://127.0.0.1:{self.demo_port}/V2/#/staffManagement"
        return raw_url

    def _wait_for_spa_settlement(self, page: Page):
        """
        Fast settlement check:
        Never waits for networkidle on #/feed.
        Only waits up to 5 seconds for #/staffManagement's actual UI element (button:has-text("Reports"))
        to become visible, then proceeds immediately to Task 1.
        """
        try:
            reports_btn = page.locator('button:has-text("Reports"), #btnReports').first
            reports_btn.wait_for(state="visible", timeout=5000)
            self.logger.info("[ENGINE] Confirmed #/staffManagement UI element (button:has-text('Reports')) is visible.")
            return
        except Exception:
            pass

        # If on add page or other view, check for relevant UI element up to 2s
        try:
            alt_elem = page.locator('button:has-text("Download Staff Data"), input[name="employeeId"], #btnCancel').first
            alt_elem.wait_for(state="visible", timeout=2000)
        except Exception:
            pass

    def _ensure_user_type_selected(self, page: Page) -> bool:
        """
        Handles DIGII's USER TYPE * checkbox multi-select dropdown with popover and OK button (image_4b5d4c.jpg):
        1. Always clicks the USER TYPE * dropdown trigger (the input/box displaying 'All User Type' directly below USER TYPE *).
        2. Waits 0.5s for the popover containing Search..., Select all, administrator, faculty, and the red OK button to open.
        3. Clicks the 'Select all' checkbox/label inside that dropdown popover (or 'faculty' as fallback).
        4. Clicks the red 'OK' button inside the popover to commit the selection.
        Also supports standard <select> elements for demo mode / fallback.
        """
        self.logger.info("[MODAL STEP 2] Handling USER TYPE * dropdown selection...")

        # Native <select> fallback (for demo mode or standard selects)
        try:
            native_select = page.locator('select#userTypeSelect:visible, select[name*="userType" i]:visible, select[name*="user_type" i]:visible')
            if native_select.count() > 0 and native_select.first.is_visible():
                for opt_lbl in ["All", "Faculty", "all", "faculty"]:
                    try:
                        native_select.first.select_option(label=opt_lbl)
                        self.logger.info(f"[MODAL STEP 2] Selected native USER TYPE option: {opt_lbl}")
                        return True
                    except Exception:
                        pass
                native_select.first.select_option(index=1)
                self.logger.info("[MODAL STEP 2] Selected first native option by index.")
                return True
        except Exception:
            pass

        # 1. Click USER TYPE * dropdown trigger (the input/box displaying "All User Type" directly below USER TYPE *)
        trigger_candidates = [
            '#userTypeTrigger',
            'xpath=//*[normalize-space(text())="USER TYPE *"]/following::*[normalize-space(text())="All User Type"][1]',
            'xpath=//*[contains(text(), "USER TYPE")]/following::div[contains(., "All User Type")][1]',
            'xpath=//*[contains(text(), "USER TYPE")]/following::input[1]',
            'xpath=//*[contains(text(), "USER TYPE")]/following::*[contains(@class, "dropdown") or contains(@class, "select") or @role="combobox"][1]',
            'text="All User Type"',
            'input[value*="All User Type" i]',
            'input[placeholder*="All User Type" i]',
            'div:has-text("All User Type")',
            'span:has-text("All User Type")',
            'div:has-text("USER TYPE *") [role="combobox"]',
            'label:has-text("USER TYPE *") ~ *',
            '[aria-label*="User Type" i]',
        ]

        trigger_clicked = False
        for sel in trigger_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    self.logger.info(f"[MODAL STEP 2] Clicking USER TYPE * dropdown trigger via [{sel}]...")
                    loc.first.click()
                    trigger_clicked = True
                    break
            except Exception:
                continue

        if not trigger_clicked:
            try:
                cand = page.get_by_text("All User Type", exact=True).first
                if cand.count() > 0 and cand.is_visible():
                    cand.click()
                    trigger_clicked = True
                    self.logger.info("[MODAL STEP 2] Clicked USER TYPE * trigger via exact text leaf.")
            except Exception:
                pass

        if not trigger_clicked:
            self.logger.warning("[MODAL STEP 2] Could not click USER TYPE * trigger directly; checking if popover is already visible.")

        # 2. Wait 0.5s for the popover containing Search..., Select all, administrator, faculty, and red OK button to open
        time.sleep(0.5)

        # 3. Click 'Select all' checkbox/label inside dropdown popover (or 'faculty' as fallback)
        select_all_clicked = False
        select_all_candidates = [
            '#userTypeSelectAll',
            'label:has-text("Select all")',
            'text="Select all"',
            'xpath=//*[normalize-space(text())="Select all"]',
            ':has-text("Select all")',
        ]
        for s_sel in select_all_candidates:
            try:
                loc = page.locator(s_sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.click()
                    select_all_clicked = True
                    self.logger.info(f"[MODAL STEP 2] Clicked 'Select all' checkbox/label via [{s_sel}]")
                    break
            except Exception:
                continue

        if not select_all_clicked:
            # Fallback to 'faculty'
            self.logger.info("[MODAL STEP 2] 'Select all' not detected; falling back to 'faculty' option...")
            faculty_candidates = [
                '#userTypeFaculty',
                'label:has-text("faculty")',
                'label:has-text("Faculty")',
                'text="faculty"',
                'text="Faculty"',
                'xpath=//*[normalize-space(text())="faculty"]',
                ':has-text("faculty")',
            ]
            for f_sel in faculty_candidates:
                try:
                    loc = page.locator(f_sel)
                    if loc.count() > 0 and loc.first.is_visible():
                        loc.first.click()
                        select_all_clicked = True
                        self.logger.info(f"[MODAL STEP 2] Clicked 'faculty' fallback option via [{f_sel}]")
                        break
                except Exception:
                    continue

        time.sleep(0.3)

        # 4. Click the red 'OK' button inside the popover to commit selection
        ok_candidates = [
            'button:has-text("OK")',
            'button:has-text("Ok")',
            'button.btn-danger:has-text("OK")',
            'xpath=//button[normalize-space(text())="OK"]',
            'button:has-text("Apply")',
        ]
        ok_clicked = False
        for ok_sel in ok_candidates:
            try:
                loc = page.locator(ok_sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.click()
                    ok_clicked = True
                    self.logger.info(f"[MODAL STEP 2] Clicked red 'OK' button inside popover via [{ok_sel}]")
                    break
            except Exception:
                continue

        if not ok_clicked:
            self.logger.warning("[MODAL STEP 2] 'OK' button in dropdown popover not found or not visible.")

        time.sleep(0.4)
        return select_all_clicked and ok_clicked

    def _handle_report_modal_download(self, page: Page, task_name: str, action_target: str) -> str:
        """
        Executes 3-step async report download modal:
        Step 1: Click 'Reports' button -> locate exact text leaf node (no broad div/span!) -> click directly -> verify header changes.
        Step 2: Check & select USER TYPE * if required/empty -> click XLSX toggle -> click first red Download XLSX button (no expect_download).
        Step 3: Wait up to 45s for 'Your Report is Ready' modal, locate final download button/link,
                expect download, save file into raw_downloads/, take screenshot proof, and close modal.
        """
        if "->" in action_target:
            trigger_sel, report_target = [s.strip() for s in action_target.split("->", 1)]
        else:
            trigger_sel = 'button:has-text("Reports"), a:has-text("Reports"), #btnReports, [title*="Reports"]'
            report_target = action_target

        self.logger.info(f"[MODAL STEP 1] Opening Reports modal via locator: [{trigger_sel}]")
        modal_trigger_loc = self._find_visible_locator(page, trigger_sel)
        if not modal_trigger_loc:
            modal_trigger_loc = page.locator(trigger_sel).first
        modal_trigger_loc.click()

        time.sleep(1)
        clean_report_name = report_target.replace("text=", "").replace('"', '').strip()
        self.logger.info(f"[MODAL STEP 1] Locating exact text leaf node for target report: [{clean_report_name}]")

        leaf_loc = None
        # 1. Exact text leaf node via get_by_text(clean_report_name, exact=True)
        try:
            cand = page.get_by_text(clean_report_name, exact=True).first
            if cand.count() > 0 and cand.is_visible():
                leaf_loc = cand
        except Exception:
            pass

        # 2. XPath leaf node with normalize-space()
        if not leaf_loc:
            try:
                cand = page.locator(f'xpath=//*[normalize-space(text())="{clean_report_name}"]').first
                if cand.count() > 0 and cand.is_visible():
                    leaf_loc = cand
            except Exception:
                pass

        # 3. Fallback to specific non-container selectors (NO broad div/span!)
        if not leaf_loc:
            specific_candidates = [
                f'.report-item:has-text("{clean_report_name}")',
                f'button:has-text("{clean_report_name}")',
                f'a:has-text("{clean_report_name}")',
                f'tr:has-text("{clean_report_name}")',
                f'text="{clean_report_name}"',
            ]
            for s_sel in specific_candidates:
                try:
                    c = page.locator(s_sel).first
                    if c.count() > 0 and c.is_visible():
                        leaf_loc = c
                        break
                except Exception:
                    continue

        if not leaf_loc:
            raise RuntimeError(f"Could not locate visible report leaf node in modal matching: {report_target}")

        self.logger.info(f"[MODAL STEP 1] Scrolling into view and clicking exact report row: [{clean_report_name}]")
        leaf_loc.scroll_into_view_if_needed()
        leaf_loc.click()

        # Verify after clicking that the modal header changes to {clean_report_name} before proceeding
        self.logger.info(f"[MODAL STEP 1] Verifying modal header updates to [{clean_report_name}]...")
        header_verified = False
        start_verify = time.time()
        while time.time() - start_verify < 6:
            header_loc = page.locator(
                f'.modal-title:has-text("{clean_report_name}"), '
                f'.modal-header:has-text("{clean_report_name}"), '
                f'h1:has-text("{clean_report_name}"), '
                f'h2:has-text("{clean_report_name}"), '
                f'h3:has-text("{clean_report_name}"), '
                f'h4:has-text("{clean_report_name}")'
            )
            if header_loc.count() > 0 and header_loc.first.is_visible():
                header_verified = True
                self.logger.info(f"[MODAL STEP 1] Verified modal header changed to [{clean_report_name}]")
                break

            modal_headers = page.locator('.modal h1, .modal h2, .modal h3, .modal h4, .modal-header, [role="dialog"] h1, [role="dialog"] h2, [role="dialog"] h3, .modal-title')
            for i in range(modal_headers.count()):
                try:
                    if clean_report_name.lower() in modal_headers.nth(i).inner_text().lower():
                        header_verified = True
                        break
                except Exception:
                    pass
            if header_verified:
                self.logger.info(f"[MODAL STEP 1] Verified modal header contains [{clean_report_name}]")
                break
            time.sleep(0.5)

        if not header_verified:
            self.logger.warning(f"[MODAL STEP 1] Modal header change to [{clean_report_name}] not explicitly detected within 6s; proceeding to Step 2.")

        # Step 2: Handle filter/format view ("Apply filters and select columns, then download")
        self.logger.info("[MODAL STEP 2] Waiting 1.5s for filter/format view ('Apply filters and select columns, then download')...")
        time.sleep(1.5)

        # Check & handle USER TYPE * required validation or empty/unselected state
        self._ensure_user_type_selected(page)

        self.logger.info("[MODAL STEP 2] Clicking XLSX format toggle (button:has-text(\"XLSX\"))...")
        xlsx_btn = self._find_visible_locator(page, 'button:has-text("XLSX")')
        if xlsx_btn:
            xlsx_btn.click()
        else:
            try:
                page.locator('button:has-text("XLSX")').first.click(timeout=5000)
            except Exception as ex:
                self.logger.warning(f"[MODAL STEP 2] XLSX toggle click warning: {ex}")

        time.sleep(1)

        self.logger.info("[MODAL STEP 2] Clicking initial red 'Download XLSX' button (to trigger Generating Report / Ready modal, no expect_download)...")
        first_dl_candidates = [
            'button:has-text("Download XLSX")',
            'button.btn-danger:has-text("Download XLSX")',
            'button.btn-primary:has-text("Download XLSX")',
            'button:has-text("Download")',
        ]
        first_dl_btn = None
        for f_sel in first_dl_candidates:
            try:
                loc = page.locator(f_sel)
                if loc.count() > 0 and loc.first.is_visible():
                    first_dl_btn = loc.first
                    break
            except Exception:
                continue

        if first_dl_btn:
            first_dl_btn.click()
        else:
            page.locator('button:has-text("Download XLSX")').first.click()

        time.sleep(1)

        # If User Type is required validation message appeared after download click, select it and re-click
        user_type_err = page.locator('text="User Type is required", text="user type is required", :has-text("User Type is required")')
        if user_type_err.count() > 0 and user_type_err.first.is_visible():
            self.logger.info("[MODAL STEP 2] Validation error appeared after download click; resolving USER TYPE * and re-clicking...")
            self._ensure_user_type_selected(page)
            time.sleep(0.5)
            if xlsx_btn:
                try:
                    xlsx_btn.click()
                except Exception:
                    pass
            if first_dl_btn:
                first_dl_btn.click()
            else:
                page.locator('button:has-text("Download XLSX")').first.click()

        # Step 3: Wait up to 45 seconds for "Generating Report / Your Report is Ready!" modal to appear
        self.logger.info("[MODAL STEP 3] Waiting up to 45s for 'Generating Report / Your Report is Ready!' modal...")
        ready_selectors = [
            'text="Your report is ready!"',
            'text="Your Report is Ready"',
            'text="Your report is ready"',
            'text="Generating Report"',
            ':has-text("Your report is ready")',
            ':has-text("Your Report is Ready")',
            ':has-text("Generating Report")',
            '.report-ready',
        ]
        ready_found = False
        start_wait = time.time()
        while time.time() - start_wait < 45:
            for r_sel in ready_selectors:
                try:
                    r_loc = page.locator(r_sel)
                    if r_loc.count() > 0 and r_loc.first.is_visible():
                        ready_found = True
                        self.logger.info(f"[MODAL STEP 3] Ready modal indicator detected via [{r_sel}] in {time.time() - start_wait:.1f}s")
                        break
                except Exception:
                    pass
            if ready_found:
                break
            time.sleep(1)

        if not ready_found:
            self.logger.warning("[MODAL STEP 3] Ready text not explicitly visible within 45s; proceeding to locate second download button.")

        # Locate the top-most visible Download XLSX button/link inside that ready modal using .last
        self.logger.info("[MODAL STEP 3] Locating visible Download XLSX button inside ready modal via .last...")
        ready_btn = page.locator('a:has-text("Download XLSX"):visible, button:has-text("Download XLSX"):visible').last
        try:
            ready_btn.wait_for(state="visible", timeout=10000)
        except Exception as ex:
            self.logger.warning(f"[MODAL STEP 3] ready_btn wait_for warning: {ex}. Checking fallback candidates...")
            fallback_btn = page.locator('a:has-text("Download"):visible, button:has-text("Download"):visible').last
            if fallback_btn.count() > 0 and fallback_btn.is_visible():
                ready_btn = fallback_btn

        self.logger.info("[MODAL STEP 3] Triggering final download from second Download XLSX button inside ready modal...")
        with page.expect_download(timeout=30000) as download_info:
            ready_btn.click()

        download = download_info.value
        clean_filename = f"{task_name}_{download.suggested_filename}"
        save_path = self.archive.downloads_dir / clean_filename
        download.save_as(str(save_path))
        self.logger.info(f"[DOWNLOAD] Saved exported modal report: {save_path.name}")

        # Capture full-page screenshot proof
        ss_filename = f"{task_name}_success.png"
        ss_full_path = self.archive.screenshots_dir / ss_filename
        try:
            page.screenshot(path=str(ss_full_path), full_page=True)
            self.logger.info(f"[PROOF] Screenshot proof captured: screenshots/{ss_filename}")
        except Exception as ss_err:
            self.logger.warning(f"Could not capture success screenshot: {ss_err}")

        # Click Close (button:has-text("Close"):visible or ×) to dismiss modal cleanly before next task
        try:
            close_btn = page.locator('button:has-text("Close"):visible, button:has-text("close"):visible, a:has-text("Close"):visible, button:has-text("×"):visible, .modal-header button:visible').last
            if close_btn.count() > 0 and close_btn.is_visible():
                close_btn.click()
                self.logger.info("[MODAL STEP 3] Closed ready modal.")
            else:
                page.keyboard.press("Escape")
        except Exception:
            pass

        return f"raw_downloads/{clean_filename}"

    def _handle_modal_email_dispatch(self, page: Page, task_name: str, action_target: str) -> Tuple[str, str]:
        """
        Handles MODAL_EMAIL_DISPATCH mode (e.g. for Staff Master Data):
        1. Clicks 'Download Staff Data' on #/staffManagement and waits ~1.5s for modal to appear.
        2. Locates and checks 'Select All' checkbox inside visible modal.
        3. Clicks bottom 'Download' button inside visible modal using .last.
        4. Waits for confirmation toast ('receive this report shortly' / .toast-success) or modal close.
        5. Captures Download_Staff_Data_Master_success.png as proof.
        6. Returns ('Queued via Email', screenshot_proof_rel).
        """
        trigger_sel = action_target or 'button:has-text("Download Staff Data"), button:has-text("Download Staff")'
        self.logger.info(f"[EMAIL_DISPATCH] Opening Download Staff Data modal via [{trigger_sel}]...")
        trigger_loc = self._find_visible_locator(page, trigger_sel)
        if not trigger_loc:
            trigger_loc = page.locator(trigger_sel).first
        trigger_loc.click()

        self.logger.info("[EMAIL_DISPATCH] Waiting 1.5s for 'Download Staff Data' modal to appear...")
        time.sleep(1.5)

        # 1. Locate and click 'Select All' checkbox inside visible modal
        self.logger.info("[EMAIL_DISPATCH] Locating and verifying 'Select All' checkbox inside visible modal...")
        select_all_candidates = [
            '.modal:visible label:has-text("Select All"):visible',
            '[role="dialog"]:visible label:has-text("Select All"):visible',
            'label:has-text("Select All"):visible',
            'label:has-text("Select all"):visible',
            'text=/Select All/i',
            '#selectAllCheckbox:visible',
            'input[type="checkbox"][id*="selectAll" i]:visible',
            'input[type="checkbox"][name*="selectAll" i]:visible',
            '.modal:visible input[type="checkbox"]:visible',
            '[role="dialog"]:visible input[type="checkbox"]:visible',
        ]

        select_all_elem = None
        for sel in select_all_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    select_all_elem = loc.first
                    break
            except Exception:
                continue

        if select_all_elem:
            try:
                select_all_elem.scroll_into_view_if_needed()
                time.sleep(0.2)

                # Check if it's already an input[type="checkbox"]
                is_input_cb = select_all_elem.evaluate("el => el.tagName.toLowerCase() === 'input' && el.type === 'checkbox'")
                if is_input_cb:
                    if not select_all_elem.is_checked():
                        select_all_elem.check()
                        self.logger.info("[EMAIL_DISPATCH] Checked 'Select All' checkbox input directly.")
                    else:
                        self.logger.info("[EMAIL_DISPATCH] 'Select All' checkbox input is already checked.")
                else:
                    # Look for input[type="checkbox"] inside or associated with label
                    inner_cb = select_all_elem.locator('input[type="checkbox"]')
                    if inner_cb.count() > 0 and inner_cb.first.is_visible():
                        if not inner_cb.first.is_checked():
                            inner_cb.first.check()
                            self.logger.info("[EMAIL_DISPATCH] Checked nested 'Select All' checkbox.")
                        else:
                            self.logger.info("[EMAIL_DISPATCH] Nested 'Select All' checkbox is already checked.")
                    else:
                        select_all_elem.click()
                        self.logger.info(f"[EMAIL_DISPATCH] Clicked 'Select All' element [{select_all_elem}]")
            except Exception as ex:
                self.logger.warning(f"[EMAIL_DISPATCH] Select All interaction warning: {ex}. Proceeding...")
        else:
            self.logger.warning("[EMAIL_DISPATCH] Could not locate visible 'Select All' checkbox, proceeding to download button...")

        time.sleep(0.8)

        # 2. Locate bottom 'Download' button inside visible modal using .last, scroll into view, and click it
        self.logger.info("[EMAIL_DISPATCH] Locating bottom 'Download' button inside visible modal using .last...")
        download_btn_candidates = [
            '.modal:visible button:has-text("Download"):visible',
            '[role="dialog"]:visible button:has-text("Download"):visible',
            'button:has-text("Download"):visible',
        ]
        download_btn = None
        for d_sel in download_btn_candidates:
            try:
                loc = page.locator(d_sel)
                if loc.count() > 0 and loc.last.is_visible():
                    download_btn = loc.last
                    break
            except Exception:
                continue

        if not download_btn:
            download_btn = page.locator('button:has-text("Download"):visible').last

        try:
            download_btn.scroll_into_view_if_needed()
            time.sleep(0.2)
            download_btn.click()
            self.logger.info("[EMAIL_DISPATCH] Clicked bottom 'Download' button in modal.")
        except Exception as dl_err:
            self.logger.warning(f"[EMAIL_DISPATCH] Failed clicking bottom download button: {dl_err}")
            page.locator('button:has-text("Download")').last.click()

        # 3. Wait for confirmation toast (receive this report shortly / .toast-success) or ~2s for modal to close
        self.logger.info("[EMAIL_DISPATCH] Waiting for confirmation toast ('receive this report shortly' / .toast-success) or modal close...")
        toast_candidates = [
            'text=/receive this report shortly/i',
            '.toast-success:visible',
            'text="You will receive this report shortly to your email address."',
            '.toast-success',
            'div:has-text("receive this report shortly")',
            ':has-text("receive this report shortly")',
            '.toast:visible',
            '.alert-success:visible',
            '.ant-message-success',
            '.notification-success',
        ]
        toast_found = False
        start_wait = time.time()
        while time.time() - start_wait < 6:
            for t_sel in toast_candidates:
                try:
                    t_loc = page.locator(t_sel)
                    if t_loc.count() > 0 and t_loc.first.is_visible():
                        toast_found = True
                        self.logger.info(f"[EMAIL_DISPATCH] Confirmation toast detected via [{t_sel}] in {time.time() - start_wait:.1f}s")
                        break
                except Exception:
                    continue
            if toast_found:
                break
            time.sleep(0.5)

        if not toast_found:
            self.logger.info("[EMAIL_DISPATCH] Toast not explicitly detected; waiting 2s for modal to close and state to settle...")
            time.sleep(2)

        # 4. Capture Download_Staff_Data_Master_success.png proof
        ss_filename = f"{task_name}_success.png"
        ss_full_path = self.archive.screenshots_dir / ss_filename
        page.screenshot(path=str(ss_full_path), full_page=True)
        screenshot_proof_rel = f"screenshots/{ss_filename}"
        self.logger.info(f"[PROOF] Screenshot proof captured: {screenshot_proof_rel}")

        downloaded_file_rel = "Queued via Email"
        return downloaded_file_rel, screenshot_proof_rel

    def _handle_direct_download(self, page: Page, task_name: str, action_target: str) -> str:
        """Locates the direct export button, initiates download, and saves the file."""
        if not action_target:
            action_target = 'button:has-text("Download Staff Data"), button:has-text("Download Staff"), button:has-text("Download")'

        self.logger.info(f"[DIRECT] Locating download action button: [{action_target}]")
        loc = self._find_visible_locator(page, action_target)
        if not loc:
            raise RuntimeError(f"Could not locate visible download button matching: {action_target}")

        self.logger.info("[DIRECT] Triggering download click...")
        with page.expect_download(timeout=self.config.download_timeout_ms) as download_info:
            loc.click()

        download = download_info.value
        clean_filename = f"{task_name}_{download.suggested_filename}"
        save_path = self.archive.downloads_dir / clean_filename
        download.save_as(str(save_path))
        self.logger.info(f"[DOWNLOAD] Saved exported file: {save_path.name}")

        return f"raw_downloads/{clean_filename}"

    def _navigate_to_add_staff_via_ui(self, page: Page):
        """
        Navigates from #/staffManagement to Add Staff form via native UI clicks:
        1. Clicks 'Actions' dropdown button on #/staffManagement (button:has-text("Actions"):visible, div:has-text("Actions"):visible).
        2. Clicks 'Add Staff' (page.get_by_text("Add Staff", exact=True).last).
        3. Waits for form elements to become visible.
        Zero raw page.goto().
        """
        if "staffmanagement/add" in page.url.lower():
            self.logger.info("[FORM DRY-RUN] Already on Staff Add page.")
            return

        # 1. Click Actions dropdown button on #/staffManagement
        self.logger.info("[FORM DRY-RUN] Locating and clicking 'Actions' dropdown button on #/staffManagement...")
        actions_candidates = [
            'button:has-text("Actions"):visible',
            'div:has-text("Actions"):visible',
            '#btnActions:visible',
            'button:has-text("Action"):visible',
            '[aria-label*="Actions" i]:visible',
        ]
        actions_loc = None
        for sel in actions_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.last.is_visible():
                    actions_loc = loc.last
                    break
            except Exception:
                continue

        if actions_loc:
            actions_loc.click()
            self.logger.info("[FORM DRY-RUN] Clicked 'Actions' dropdown button.")
            time.sleep(0.5)
        else:
            self.logger.warning("[FORM DRY-RUN] 'Actions' button not found; checking for direct 'Add Staff' element...")

        # 2. Click 'Add Staff' (page.get_by_text("Add Staff", exact=True).last)
        self.logger.info("[FORM DRY-RUN] Locating and clicking 'Add Staff'...")
        add_staff_loc = None
        try:
            cand = page.get_by_text("Add Staff", exact=True).last
            if cand.count() > 0 and cand.is_visible():
                add_staff_loc = cand
        except Exception:
            pass

        if not add_staff_loc:
            add_candidates = [
                'a:has-text("Add Staff"):visible',
                'button:has-text("Add Staff"):visible',
                'div:has-text("Add Staff"):visible',
                ':has-text("Add Staff"):visible',
            ]
            for sel in add_candidates:
                try:
                    c = page.locator(sel)
                    if c.count() > 0 and c.last.is_visible():
                        add_staff_loc = c.last
                        break
                except Exception:
                    continue

        if add_staff_loc:
            add_staff_loc.click()
            self.logger.info("[FORM DRY-RUN] Clicked 'Add Staff'.")
        else:
            self.logger.warning("[FORM DRY-RUN] 'Add Staff' element not found via UI.")

        # Wait for form elements to become visible
        try:
            form_indicator = page.locator('#employeeId, input[name="employeeId"], #addStaffForm, input[name="firstName"]').first
            form_indicator.wait_for(state="visible", timeout=10000)
            self.logger.info("[FORM DRY-RUN] Add Staff form loaded successfully via native UI navigation.")
        except Exception as ex:
            self.logger.warning(f"[FORM DRY-RUN] Add Staff form wait warning: {ex}")

    def _handle_form_dry_run(self, page: Page, task_name: str, action_target: str) -> str:
        """
        Safely populates New Staff Form (#/staffManagement/add):
        - Navigates via Actions -> Add Staff if not already on add form
        - Aggressively clears autofill (Employee ID & Password)
        - Fills safe demo values
        - Captures 04_Form_Dry_Run_Populated.png proof
        - STRICT SAFETY: Clicks CANCEL. Zero database writes.
        """
        self.logger.info("[FORM DRY-RUN] Commencing safe form entry on Staff Management Add page...")

        if "staffmanagement/add" not in page.url.lower():
            self._navigate_to_add_staff_via_ui(page)

        # 1. Employee ID
        emp_selectors = [
            'input[name="employeeId"]',
            '#employeeId',
            'input[placeholder*="Employee ID" i]',
            'input[formcontrolname*="employeeId" i]',
            'input[name*="emp" i]',
        ]
        self._safe_clear_and_fill(page, emp_selectors, "DEMO_TEST_001", "Employee ID")

        # 2. User Type: Faculty
        user_type_selectors = [
            'select[name="userType"]',
            '#userType',
            'select[formcontrolname*="userType" i]',
            '[formcontrolname*="userType" i]',
        ]
        self._safe_select_option(page, user_type_selectors, "Faculty", "User Type")

        # 3. Department: Select first available option
        dept_selectors = [
            'select[name="department"]',
            '#department',
            'select[formcontrolname*="department" i]',
            '[formcontrolname*="department" i]',
        ]
        self._safe_select_first_available(page, dept_selectors, "Department")

        # 4. First Name: Demo
        fn_selectors = [
            'input[name="firstName"]',
            '#firstName',
            'input[placeholder*="First Name" i]',
            'input[formcontrolname*="firstName" i]',
        ]
        self._safe_clear_and_fill(page, fn_selectors, "Demo", "First Name")

        # 5. Last name: Faculty
        ln_selectors = [
            'input[name="lastName"]',
            '#lastName',
            'input[placeholder*="Last Name" i]',
            'input[placeholder*="Last name" i]',
            'input[formcontrolname*="lastName" i]',
        ]
        self._safe_clear_and_fill(page, ln_selectors, "Faculty", "Last name")

        # 6. Gender: Male
        gender_selectors = [
            'input[name="gender"][value="Male"]',
            '#genderMale',
            'label:has-text("Male")',
            'input[type="radio"][value*="Male" i]',
            'select[name="gender"]',
        ]
        self._safe_click_or_select(page, gender_selectors, "Male", "Gender")

        # 7. Email ID: demo.test@poornima.org
        email_selectors = [
            'input[type="email"]',
            'input[name="email"]',
            '#email',
            'input[placeholder*="Email" i]',
            'input[formcontrolname*="email" i]',
        ]
        self._safe_clear_and_fill(page, email_selectors, "demo.test@poornima.org", "Email ID")

        # 8. Phone Number: 9999999999
        phone_selectors = [
            'input[type="tel"]',
            'input[name="phoneNumber"]',
            '#phoneNumber',
            'input[placeholder*="Phone" i]',
            'input[formcontrolname*="phoneNumber" i]',
        ]
        self._safe_clear_and_fill(page, phone_selectors, "9999999999", "Phone Number")

        # 9. Password: DemoPass@123
        pwd_selectors = [
            'input[type="password"]',
            '#password',
            'input[name="password"]',
            'input[placeholder*="Password" i]',
            'input[formcontrolname*="password" i]',
        ]
        self._safe_clear_and_fill(page, pwd_selectors, "DemoPass@123", "Password")

        time.sleep(1)

        # 10. Capture Proof Screenshot
        proof_filename = "04_Form_Dry_Run_Populated.png"
        proof_path = self.archive.screenshots_dir / proof_filename
        page.screenshot(path=str(proof_path), full_page=True)
        self.logger.info(f"[PROOF] Populated form screenshot saved: {proof_filename}")

        # 11. CRITICAL SAFETY RULE: Click CANCEL (Never Create User)
        self.logger.info("[SAFETY CHECK] Enforcing zero database writes. Finding 'Cancel' button...")
        cancel_selectors = [
            '#btnCancel',
            'button:has-text("Cancel")',
            'a:has-text("Cancel")',
            'button[type="button"]:has-text("Cancel")',
            '[aria-label*="Cancel"]',
        ]
        cancel_loc = self._find_visible_locator(page, ",".join(cancel_selectors))
        if cancel_loc:
            self.logger.info("[SAFETY ACTION] Clicking 'Cancel' button. ZERO records created.")
            cancel_loc.click()
            time.sleep(1)
        else:
            self.logger.warning("[SAFETY WARNING] Could not find 'Cancel' button. Navigating away safely.")
            page.goto(self._resolve_target_url("https://poornima.digiicampus.com/V2/#/staffManagement"))

        return f"screenshots/{proof_filename}"

    def _safe_clear_and_fill(self, page: Page, selector_candidates: List[str], value: str, field_name: str):
        """Clicks, aggressively clears autofill, and fills the text input."""
        for sel in selector_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    self.logger.info(f"[FORM] Clearing autofill & filling {field_name} -> [{sel}]")
                    loc.first.click()
                    try:
                        loc.first.clear()
                    except Exception:
                        pass
                    loc.first.press("Control+A")
                    loc.first.press("Backspace")
                    time.sleep(0.05)
                    loc.first.fill(value)
                    return True
            except Exception:
                continue
        self.logger.warning(f"[FORM] Warning: Field '{field_name}' not matched with candidates.")
        return False

    def _safe_select_option(self, page: Page, selector_candidates: List[str], label_or_val: str, field_name: str):
        """Selects option from standard <select> or custom dropdown."""
        for sel in selector_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    tag = loc.first.evaluate("el => el.tagName.toLowerCase()")
                    if tag == "select":
                        loc.first.select_option(label=label_or_val)
                    else:
                        loc.first.click()
                        time.sleep(0.3)
                        opt = page.locator(f'mat-option:has-text("{label_or_val}"), [role="option"]:has-text("{label_or_val}"), .dropdown-item:has-text("{label_or_val}"), text="{label_or_val}"')
                        if opt.count() > 0:
                            opt.first.click()
                    self.logger.info(f"[FORM] Selected '{label_or_val}' for {field_name}")
                    return True
            except Exception:
                continue
        self.logger.warning(f"[FORM] Warning: Could not select '{label_or_val}' for {field_name}")
        return False

    def _safe_select_first_available(self, page: Page, selector_candidates: List[str], field_name: str):
        """Selects first non-empty option in dropdown."""
        for sel in selector_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    tag = loc.first.evaluate("el => el.tagName.toLowerCase()")
                    if tag == "select":
                        loc.first.select_option(index=1)
                    else:
                        loc.first.click()
                        time.sleep(0.3)
                        opt = page.locator('mat-option, [role="option"], .dropdown-item, li')
                        if opt.count() > 1:
                            opt.nth(1).click()
                        elif opt.count() > 0:
                            opt.first.click()
                    self.logger.info(f"[FORM] Selected first available option for {field_name}")
                    return True
            except Exception:
                continue
        self.logger.warning(f"[FORM] Warning: Could not select first available option for {field_name}")
        return False

    def _safe_click_or_select(self, page: Page, selector_candidates: List[str], val: str, field_name: str):
        """Handles radio buttons, toggle buttons, or select for fields like Gender."""
        for sel in selector_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    tag = loc.first.evaluate("el => el.tagName.toLowerCase()")
                    if tag == "select":
                        loc.first.select_option(label=val)
                    else:
                        loc.first.click()
                    self.logger.info(f"[FORM] Selected '{val}' for {field_name}")
                    return True
            except Exception:
                continue
        self.logger.warning(f"[FORM] Warning: Could not set '{val}' for {field_name}")
        return False

    def _find_visible_locator(self, page: Page, comma_separated_selectors: str):
        """Finds the first visible Playwright locator from candidates."""
        candidates = [s.strip() for s in comma_separated_selectors.split(",") if s.strip()]
        for sel in candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    return loc.first
            except Exception:
                continue
        return None


# -----------------------------------------------------------------------------
# 6. Executive Excel Audit Report Generator
# -----------------------------------------------------------------------------
class AuditReportGenerator:
    """Creates a beautifully styled, clickable Executive Excel Audit Report using openpyxl."""

    HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    SUCCESS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    SUCCESS_FONT = Font(name="Calibri", size=10, bold=True, color="006100")

    FAILED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    FAILED_FONT = Font(name="Calibri", size=10, bold=True, color="9C0006")

    HYPERLINK_FONT = Font(name="Calibri", size=10, color="0563C1", underline="single")
    DATA_FONT = Font(name="Calibri", size=10)

    THIN_BORDER = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    def __init__(self, archive: RunArchiveManager, logger: logging.Logger):
        self.archive = archive
        self.logger = logger

    def generate(self, records: List[Dict[str, Any]]) -> Path:
        """Builds and saves Audit_Report.xlsx."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Executive_Audit_Report"

        headers = [
            "Timestamp",
            "Task_Name",
            "Target_URL",
            "Status",
            "Downloaded_File",
            "Screenshot_Proof",
            "Error_Notes",
        ]
        ws.append(headers)

        ws.row_dimensions[1].height = 28
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = self.HEADER_FILL
            cell.font = self.HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.THIN_BORDER

        for row_idx, rec in enumerate(records, start=2):
            ws.row_dimensions[row_idx].height = 22

            # 1. Timestamp
            c1 = ws.cell(row=row_idx, column=1, value=rec["Timestamp"])
            c1.alignment = Alignment(horizontal="center", vertical="center")
            c1.font = self.DATA_FONT
            c1.border = self.THIN_BORDER

            # 2. Task Name
            c2 = ws.cell(row=row_idx, column=2, value=rec["Task_Name"])
            c2.alignment = Alignment(horizontal="left", vertical="center")
            c2.font = self.DATA_FONT
            c2.border = self.THIN_BORDER

            # 3. Target URL
            c3 = ws.cell(row=row_idx, column=3, value=rec["Target_URL"])
            c3.alignment = Alignment(horizontal="left", vertical="center")
            c3.font = self.DATA_FONT
            c3.border = self.THIN_BORDER

            # 4. Status Badge
            status_val = rec["Status"]
            c4 = ws.cell(row=row_idx, column=4, value=status_val)
            c4.alignment = Alignment(horizontal="center", vertical="center")
            c4.border = self.THIN_BORDER
            if "SUCCESS" in status_val:
                c4.fill = self.SUCCESS_FILL
                c4.font = self.SUCCESS_FONT
            else:
                c4.fill = self.FAILED_FILL
                c4.font = self.FAILED_FONT

            # 5. Downloaded File (Clickable Link)
            download_val = rec["Downloaded_File"]
            c5 = ws.cell(row=row_idx, column=5)
            c5.alignment = Alignment(horizontal="left", vertical="center")
            c5.border = self.THIN_BORDER
            if download_val and download_val != "N/A" and not download_val.startswith("N/A"):
                if download_val.startswith("Queued") or not ("/" in download_val or "\\" in download_val):
                    c5.value = download_val
                    c5.font = self.DATA_FONT
                else:
                    c5.value = os.path.basename(download_val)
                    c5.hyperlink = download_val
                    c5.font = self.HYPERLINK_FONT
            else:
                c5.value = download_val or "N/A"
                c5.font = self.DATA_FONT

            # 6. Screenshot Proof (Clickable Link)
            ss_val = rec["Screenshot_Proof"]
            c6 = ws.cell(row=row_idx, column=6)
            c6.alignment = Alignment(horizontal="left", vertical="center")
            c6.border = self.THIN_BORDER
            if ss_val and ss_val != "N/A":
                c6.value = os.path.basename(ss_val)
                c6.hyperlink = ss_val
                c6.font = self.HYPERLINK_FONT
            else:
                c6.value = "N/A"
                c6.font = self.DATA_FONT

            # 7. Error Notes
            c7 = ws.cell(row=row_idx, column=7, value=rec["Error_Notes"])
            c7.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c7.font = self.DATA_FONT
            c7.border = self.THIN_BORDER

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(min(max_len + 4, 45), 14)

        ws.freeze_panes = "A2"
        wb.save(self.archive.report_path)
        self.logger.info(f"[REPORT] Executive Audit Report generated at: {self.archive.report_path.resolve()}")
        return self.archive.report_path

    def auto_open_report(self):
        """Attempts to open the report automatically on Windows."""
        try:
            if sys.platform == "win32" and self.archive.report_path.exists():
                self.logger.info(f"[REPORT] Launching {self.archive.report_path.name} in Microsoft Excel...")
                os.startfile(str(self.archive.report_path.resolve()))
        except Exception as e:
            self.logger.warning(f"Could not auto-open Excel report: {e}")


# -----------------------------------------------------------------------------
# 7. Main Entry Point
# -----------------------------------------------------------------------------
def setup_logging(archive: RunArchiveManager) -> logging.Logger:
    logger = logging.getLogger("DigiiMondayBot")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    fh = logging.FileHandler(str(archive.log_path), encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    return logger


def main():
    parser = argparse.ArgumentParser(description="DIGII ERP Monday Automation Bot")
    parser.add_argument("--config", default="config/urls_config.xlsx", help="Path to urls_config.xlsx")
    parser.add_argument("--headless", action="store_true", help="Force headless browser execution")
    parser.add_argument("--demo", action="store_true", help="Force local interactive demo mode")
    parser.add_argument("--dry-run", action="store_true", help="Perform syntax, config and session dry run")
    args = parser.parse_args()

    archive = RunArchiveManager()
    logger = setup_logging(archive)

    logger.info("=" * 78)
    logger.info("        DIGII ERP AUTOMATION TOOL (Digii_Monday_Bot)")
    logger.info(f" Run Output Directory: {archive.run_dir.resolve()}")
    logger.info("=" * 78)

    config = BotConfig()
    if args.headless:
        config.headless = True

    demo_server = None
    demo_port = None
    if args.demo or config.is_placeholder_config():
        logger.info("[MODE] Detected demo flag or template credentials. Initiating Poornima DIGII DEMO MODE.")
        demo_server = LocalDemoServer()
        demo_port = demo_server.start()
        config.login_url = f"http://127.0.0.1:{demo_port}/V2/#/login"
        config.dashboard_url = f"http://127.0.0.1:{demo_port}/V2/#/staffManagement"
        config.user = "demo_admin"
        config.password = "demo_password"
        logger.info(f"[MODE] Local mock ERP portal running at http://127.0.0.1:{demo_port}")

    try:
        reader = ExcelConfigReader(args.config)
        tasks = reader.load_active_tasks()
        logger.info(f"[CONFIG] Loaded {len(tasks)} active task(s) from {args.config}")
        for t in tasks:
            logger.info(f"  - [{t['mode']}] {t['task_name']} -> {t['target_url']}")
    except Exception as e:
        logger.critical(f"[CONFIG FATAL] Failed reading {args.config}: {e}")
        sys.exit(1)

    if not tasks:
        logger.warning("[CONFIG] No active tasks found to run (Active=TRUE). Exiting.")
        sys.exit(0)

    try:
        with sync_playwright() as p:
            auth_mgr = AuthManager(config, archive, logger, demo_port=demo_port)
            context = auth_mgr.create_persistent_browser_context(p)

            # Ensure session is authenticated (reusing session or logging in with 45s verification wait)
            context = auth_mgr.ensure_authenticated_session(context)

            if args.dry_run:
                logger.info("[DRY-RUN] Authentication and browser context initialized successfully. Dry-run complete.")
                context.close()
                return

            engine = DigiiAutomationEngine(config, archive, logger, demo_port=demo_port)
            engine.execute_tasks(context, tasks)

            context.close()

            reporter = AuditReportGenerator(archive, logger)
            reporter.generate(engine.audit_records)
            reporter.auto_open_report()

    except Exception as e:
        logger.critical(f"[CRITICAL EXECUTION ERROR] {e}", exc_info=True)
        sys.exit(1)
    finally:
        if demo_server:
            demo_server.stop()

    logger.info("=" * 78)
    logger.info(f"[COMPLETE] Digii_Monday_Bot completed. Report: {archive.report_path}")
    logger.info("=" * 78)


if __name__ == "__main__":
    main()
