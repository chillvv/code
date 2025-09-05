    #!/usr/bin/env python3
    # -*- coding: utf-8 -*-
    """
    main.py - WeChat automation sender (Enter-only sending)
    Replaces your previous main.py. Key features:
    - Use Ctrl+K to search contacts/groups.
    - Locate input box via UIA 'Edit' or by clicking a computed coordinate in the window.
    - Only press Enter to send messages.
    - DPI aware for high-DPI systems.
    - Chunk long messages and paste via clipboard to maximize reliability.
    - Robust retries and logging.
    """

    import time
    import sys
    import logging
    import platform
    import traceback
    from typing import Optional, Tuple, List

    # Third-party libs (ensure installed: pywinauto, pyautogui, pyperclip)
    try:
        from pywinauto import Application, mouse
        from pywinauto.keyboard import send_keys as pywinauto_send_keys
    except Exception:
        Application = None
        mouse = None
        pywinauto_send_keys = None

    try:
        import pyautogui
    except Exception:
        pyautogui = None

    try:
        import pyperclip
    except Exception:
        pyperclip = None

    # Configure logging
    logger = logging.getLogger("WeChatAutoSender")
    logger.setLevel(logging.DEBUG)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    ch.setFormatter(formatter)
    logger.addHandler(ch)


    class WeChatAutoSender:
        def __init__(self, wechat_title_re: str = r"WeChat|微信", backend: str = "uia"):
            """
            wechat_title_re: regex to find WeChat main window title
            backend: pywinauto backend, default 'uia' (better for modern apps)
            """
            self.wechat_title_re = wechat_title_re
            self.backend = backend
            self.app = None
            self.main_window = None
            self._ensure_dpi_awareness()

        def _ensure_dpi_awareness(self):
            """Make process DPI aware to reduce coordinate scaling issues on high-DPI displays."""
            if platform.system().lower() != "windows":
                return
            try:
                import ctypes
                # Try both legacy and newer API if available
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception:
                    # Windows 8.1+ alternative
                    try:
                        PROCESS_PER_MONITOR_DPI_AWARE = 2
                        shcore = ctypes.windll.shcore
                        shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
                    except Exception:
                        pass
                logger.debug("Set process DPI aware (if supported).")
            except Exception as e:
                logger.debug(f"Failed to set DPI awareness: {e}")

        def connect(self, timeout: float = 5.0) -> bool:
            """Connect to running WeChat application and set self.main_window"""
            if Application is None:
                logger.error("pywinauto is required. Please install pywinauto.")
                return False
            try:
                self.app = Application(backend=self.backend).connect(title_re=self.wechat_title_re, timeout=timeout)
                # pick the best matching top-level window
                self.main_window = self.app.window(title_re=self.wechat_title_re)
                # Bring to foreground
                try:
                    self.main_window.set_focus()
                    self.main_window.set_foreground()
                except Exception:
                    try:
                        self.main_window.wrapper_object().set_focus()
                    except Exception:
                        pass
                logger.debug("Connected to WeChat window.")
                return True
            except Exception as e:
                logger.exception("Failed to connect to WeChat window.")
                return False

        def _bring_wechat_front(self):
            """Ensure WeChat window is foreground."""
            if not self.main_window:
                return
            try:
                self.main_window.set_focus()
                self.main_window.set_foreground()
            except Exception:
                try:
                    self.main_window.wrapper_object().set_focus()
                except Exception:
                    pass
            time.sleep(0.15)

        def _press_ctrl_k(self):
            """Send Ctrl+K to open search box (use pywinauto if available, fallback to pyautogui)."""
            logger.debug("Pressing Ctrl+K to open global search.")
            try:
                if pywinauto_send_keys:
                    # pywinauto expects format '^k'
                    pywinauto_send_keys("^k")
                elif pyautogui:
                    pyautogui.hotkey("ctrl", "k")
                else:
                    logger.error("No keyboard tool available (pywinauto or pyautogui).")
                time.sleep(0.35)
            except Exception:
                logger.exception("Failed to press Ctrl+K")

        def _paste_text(self, text: str):
            """Paste text via clipboard. Ensure pyperclip is installed."""
            if pyperclip is None:
                raise RuntimeError("pyperclip is required for reliable clipboard paste. Install it.")
            pyperclip.copy(text)
            time.sleep(0.06)
            # Use Ctrl+V (pywinauto or pyautogui)
            try:
                if pywinauto_send_keys:
                    pywinauto_send_keys("^v")
                elif pyautogui:
                    pyautogui.hotkey("ctrl", "v")
                else:
                    raise RuntimeError("No paste method available.")
                time.sleep(0.08)
            except Exception:
                logger.exception("Failed to send paste hotkey.")

        def _find_input_edit(self) -> Optional[object]:
            """Try to find the input Edit control via UIA. Return the control or None."""
            if not self.main_window:
                return None
            try:
                # find descendants with control_type 'Edit'
                edits = [w for w in self.main_window.descendants(control_type="Edit")]
                if edits:
                    # often the last Edit is the message input; pick last
                    logger.debug(f"Found {len(edits)} Edit controls; using last as input box.")
                    return edits[-1]
                else:
                    logger.debug("No Edit controls found in UIA tree.")
                    return None
            except Exception:
                logger.exception("Error while searching for Edit controls.")
                return None

        def _click_input_by_rect(self) -> Tuple[int, int]:
            """
            Compute a safe coordinate inside the conversation area (near bottom)
            based on the main window rectangle and click it.
            Returns (x, y) clicked.
            """
            # We'll compute roughly 35% from left and 80px above bottom as safe input area.
            try:
                rect = self.main_window.rectangle()
                left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
                width = right - left
                height = bottom - top
                x = left + int(width * 0.35)
                y = bottom - 80
                # clamp
                if x < left + 10:
                    x = left + 10
                if y < top + 40:
                    y = top + 40
                logger.debug(f"Clicking computed coords: ({x}, {y}) based on window rect {rect}.")
                # Prefer pywinauto.mouse.click since it's in same coordinate system
                if mouse:
                    mouse.click(button="left", coords=(x, y))
                elif pyautogui:
                    pyautogui.click(x, y)
                else:
                    raise RuntimeError("No mouse click method available.")
                time.sleep(0.18)
                return x, y
            except Exception:
                logger.exception("Failed to compute/click input area via window rect. Falling back to screen center bottom.")
                # Fallback: click near bottom middle of primary screen
                try:
                    screen_w, screen_h = pyautogui.size() if pyautogui else (1024, 768)
                    x = int(screen_w * 0.5)
                    y = max(50, screen_h - 120)
                    if pyautogui:
                        pyautogui.click(x, y)
                    elif mouse:
                        mouse.click(button="left", coords=(x, y))
                    time.sleep(0.18)
                    return x, y
                except Exception:
                    logger.exception("Fallback click also failed.")
                    return -1, -1

        def _ensure_focus_on_input(self) -> Optional[object]:
            """
            Ensure message input box is focused and return the Edit control if found.
            Steps:
                1. Try to find Edit via UIA
                2. If not found, click computed coordinate and try again
            """
            edit = self._find_input_edit()
            if edit:
                try:
                    edit.set_focus()
                    time.sleep(0.08)
                    return edit
                except Exception:
                    # sometimes set_focus fails; continue
                    pass

            # If not found, click the computed area and try to grab the Edit again
            self._click_input_by_rect()
            # wait a bit for UI to update
            time.sleep(0.18)
            edit = self._find_input_edit()
            if edit:
                try:
                    edit.set_focus()
                    time.sleep(0.06)
                except Exception:
                    pass
                return edit
            # final fallback: try to send a Tab key to shift focus into input area (best-effort)
            try:
                if pywinauto_send_keys:
                    pywinauto_send_keys("{TAB}")
                elif pyautogui:
                    pyautogui.press("tab")
                time.sleep(0.12)
            except Exception:
                pass
            edit = self._find_input_edit()
            return edit

        def _open_chat(self, target: str, max_attempts: int = 3) -> bool:
            """
            Open the contact/group by searching (Ctrl+K) and hitting Enter.
            Returns True if it thinks chat is opened.
            """
            for attempt in range(1, max_attempts + 1):
                try:
                    self._bring_wechat_front()
                    self._press_ctrl_k()
                    # Paste the target name and press Enter
                    if pyperclip is None:
                        raise RuntimeError("pyperclip required to paste target name.")

                    pyperclip.copy(target)
                    time.sleep(0.06)
                    # paste
                    if pywinauto_send_keys:
                        pywinauto_send_keys("^v")
                    elif pyautogui:
                        pyautogui.hotkey("ctrl", "v")
                    else:
                        raise RuntimeError("No paste method available.")

                    time.sleep(0.25)
                    # Press Enter to select
                    if pywinauto_send_keys:
                        pywinauto_send_keys("{ENTER}")
                    elif pyautogui:
                        pyautogui.press("enter")
                    else:
                        raise RuntimeError("No keyboard send method available.")

                    # allow time for chat window to switch
                    time.sleep(0.9 + attempt * 0.2)
                    logger.debug(f"Search attempt {attempt} for target '{target}' done.")
                    # quick heuristic: try finding input edit to confirm
                    edit = self._find_input_edit()
                    if edit:
                        logger.debug("Found input edit after opening chat; assume chat opened.")
                        return True
                    # else if not found, continue attempts
                except Exception:
                    logger.exception("Exception while trying to open chat.")
                    time.sleep(0.3)
            logger.warning(f"Failed to open chat '{target}' after {max_attempts} attempts.")
            return False

        def _send_via_clipboard_chunks(self, text: str, chunk_size: int = 900, pause_between_chunks: float = 0.15) -> bool:
            """
            Paste text in chunks and press Enter after each chunk to send as separate messages
            (if you want single message, keep chunk_size large; here we chunk to be robust).
            Returns True if at least one chunk was sent.
            """
            if not text:
                return False
            sent_any = False
            # Split preserving words where possible — simple chunking
            pos = 0
            n = len(text)
            while pos < n:
                chunk = text[pos: pos + chunk_size]
                # avoid splitting surrogate pairs/combining codepoints crudely — but this simple chop is OK for typical text
                try:
                    self._paste_text(chunk)
                except Exception:
                    logger.exception("Failed to paste chunk.")
                    return False
                # Press Enter to send (user requested Enter-only)
                try:
                    if pywinauto_send_keys:
                        pywinauto_send_keys("{ENTER}")
                    elif pyautogui:
                        pyautogui.press("enter")
                    else:
                        raise RuntimeError("No keyboard method to press Enter.")
                    sent_any = True
                except Exception:
                    logger.exception("Failed to press Enter after paste.")
                    return False
                pos += chunk_size
                time.sleep(pause_between_chunks)
            return sent_any

        def send_message(self, target: str, message: str, open_attempts: int = 3) -> bool:
            """
            Public API: send message to contact/group 'target'. Returns True if sent.
            """
            try:
                # Connect if needed
                if not self.main_window:
                    ok = self.connect(timeout=6)
                    if not ok:
                        logger.error("Unable to connect to WeChat. Aborting send.")
                        return False

                # Bring front and open chat
                if not self._open_chat(target, max_attempts=open_attempts):
                    logger.error("Open chat failed.")
                    # still try to ensure focus/attempt sending (best-effort)
                # Ensure input area focused
                edit = self._ensure_focus_on_input()
                # If input found, clear selection and ensure ready
                if edit:
                    try:
                        # select all then backspace to ensure clean input (non-destructive for temporary test)
                        edit.type_keys("^a{BACKSPACE}")
                    except Exception:
                        pass
                else:
                    # If no edit control, attempt clicking computed input area (already done in _ensure_focus)
                    logger.debug("Input edit not found after attempts; proceeding with click-based input focus.")

                # Now paste & send using chunked clipboard method
                sent = self._send_via_clipboard_chunks(message)
                if sent:
                    logger.info(f"Message sent to '{target}'.")
                    return True
                else:
                    logger.warning("No chunks sent (empty message or paste failed).")
                    return False
            except Exception as e:
                logger.exception("send_message exception")
                return False


    def example_run():
        """
        Example usage: run script directly to send a test message.
        Adjust 'target' and 'message' as needed.
        """
        sender = WeChatAutoSender()
        ok = sender.connect(timeout=6)
        if not ok:
            logger.error("Could not find WeChat window. Make sure WeChat is running and you are logged in.")
            return

        # Example - replace with the real contact/group name exactly as it appears in WeChat
        target = "测试群"  # <-- 改成你要发送的联系人或群名
        message = "这是来自自动化脚本的测试消息。\n如能收到则说明发送逻辑正常。"

        # Attempt send
        result = sender.send_message(target, message)
        if result:
            logger.info("Example send succeeded.")
        else:
            logger.error("Example send failed. Check logs above for clues.")


    if __name__ == "__main__":
        # If arguments are provided, allow CLI usage: python main.py "target name" "message..."
        if len(sys.argv) >= 3:
            target_arg = sys.argv[1]
            msg_arg = " ".join(sys.argv[2:])
            s = WeChatAutoSender()
            if not s.connect(timeout=6):
                logger.error("Connection failed from CLI invocation.")
                sys.exit(1)
            ok = s.send_message(target_arg, msg_arg)
            sys.exit(0 if ok else 2)
        else:
            # No args -> run example
            logger.info("Running example_run. To use from CLI: python main.py \"target\" \"message...\"")
            example_run()
