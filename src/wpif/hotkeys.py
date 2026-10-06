from __future__ import annotations

"""Global hotkey helper for smart paste.

Hotkeys only *request* a paste. The preview UI (or standalone loop) should
execute Word COM work on a suitable thread.

- Ctrl+Alt+V: request smart paste
- Ctrl+V: only while hijack is enabled; Word focused + image clipboard
"""

import threading
import time
from collections.abc import Callable

from .config import load_config
from .word_paste import clipboard_has_image, paste_and_fit, word_is_foreground

RequestHandler = Callable[[str], None]

_passthrough = False
_started = False
_ctrl_v_registered = False
_lock = threading.Lock()
_handler: RequestHandler | None = None
_pynput_listener = None


def set_request_handler(handler: RequestHandler | None) -> None:
    """Set callback invoked as handler(source) from the hotkey thread."""
    global _handler
    _handler = handler


def _emit(source: str) -> None:
    print(f"[hotkey] fired: {source}")
    if _handler is not None:
        _handler(source)
        return
    # Standalone fallback: paste here with COM init.
    outcome = paste_and_fit(load_config())
    print(f"[{source}] smart paste:", outcome)


def _on_smart_hotkey() -> None:
    _emit("ctrl+alt+v")


def _on_ctrl_v() -> None:
    global _passthrough
    import keyboard  # type: ignore

    if _passthrough:
        return

    if not load_config().hijack_ctrl_v:
        _passthrough = True
        try:
            keyboard.send("ctrl+v")
        finally:
            _passthrough = False
        return

    if not word_is_foreground() or not clipboard_has_image():
        _passthrough = True
        try:
            keyboard.send("ctrl+v")
        finally:
            _passthrough = False
        return

    _emit("ctrl+v")


def _sync_ctrl_v_registration() -> None:
    global _ctrl_v_registered
    import keyboard  # type: ignore

    want = bool(load_config().hijack_ctrl_v)
    with _lock:
        if want and not _ctrl_v_registered:
            keyboard.add_hotkey("ctrl+v", _on_ctrl_v, suppress=True)
            _ctrl_v_registered = True
            print("Ctrl+V hijack enabled (Word + image clipboard only)")
        elif not want and _ctrl_v_registered:
            try:
                keyboard.remove_hotkey("ctrl+v")
            except KeyError:
                pass
            _ctrl_v_registered = False
            print("Ctrl+V hijack disabled")


def _watch_settings() -> None:
    while True:
        try:
            _sync_ctrl_v_registration()
        except Exception as exc:  # noqa: BLE001
            print("hotkey settings sync error:", exc)
        time.sleep(0.5)


def _start_pynput_ctrl_alt_v() -> str:
    """Use pynput for Ctrl+Alt+V — usually works without admin."""
    global _pynput_listener
    from pynput import keyboard as pynput_keyboard

    _pynput_listener = pynput_keyboard.GlobalHotKeys({"<ctrl>+<alt>+v": _on_smart_hotkey})
    _pynput_listener.daemon = True
    _pynput_listener.start()
    return "pynput:Ctrl+Alt+V"


def _start_keyboard_ctrl_v_watcher() -> str:
    """keyboard lib is used only when hijack needs suppress."""
    import keyboard  # type: ignore

    # Touch the module so failures surface early.
    _ = keyboard
    threading.Thread(target=_watch_settings, name="wpif-hotkey-watch", daemon=True).start()
    _sync_ctrl_v_registration()
    return "keyboard:Ctrl+V(optional)"


def start_hotkeys(*, blocking: bool = True, on_status: Callable[[str], None] | None = None) -> None:
    global _started

    with _lock:
        if _started:
            if on_status:
                on_status("Hotkeys already running")
            return
        _started = True

    backends: list[str] = []
    errors: list[str] = []

    try:
        backends.append(_start_pynput_ctrl_alt_v())
    except Exception as exc:  # noqa: BLE001
        errors.append(f"Ctrl+Alt+V failed: {exc}")

    try:
        backends.append(_start_keyboard_ctrl_v_watcher())
    except Exception as exc:  # noqa: BLE001
        errors.append(f"Ctrl+V helper failed: {exc}")

    if backends:
        msg = "热键已启动: " + ", ".join(backends)
    else:
        msg = "热键启动失败"
    if errors:
        msg += " | " + " ; ".join(errors)

    print(msg)
    if on_status:
        on_status(msg)

    if blocking:
        # Keep process alive for standalone mode.
        if _pynput_listener is not None:
            _pynput_listener.join()
        else:
            while True:
                time.sleep(1)


def start_hotkeys_background(on_status: Callable[[str], None] | None = None) -> threading.Thread:
    thread = threading.Thread(
        target=start_hotkeys,
        kwargs={"blocking": True, "on_status": on_status},
        name="wpif-hotkeys",
        daemon=True,
    )
    thread.start()
    return thread


def main() -> None:
    print("Starting Word Paste Image Fit hotkeys. Ctrl+C to exit.")
    start_hotkeys(blocking=True)


if __name__ == "__main__":
    main()
