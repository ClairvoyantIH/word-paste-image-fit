from __future__ import annotations

"""Global hotkey helper for smart paste.

- Ctrl+Alt+V: always smart-paste into Word (when this helper is running)
- Ctrl+V: only while settings.hijack_ctrl_v is on; then only intercepts when
  Word is focused and the clipboard has an image (otherwise passes through)

Run alone:
    python -m wpif.hotkeys

Or leave the preview window open; it starts this helper in the background.
"""

import threading
import time
from typing import Callable

from .config import load_config
from .word_paste import clipboard_has_image, paste_and_fit, word_is_foreground

_passthrough = False
_started = False
_ctrl_v_registered = False
_lock = threading.Lock()


def _smart_paste(source: str) -> None:
    outcome = paste_and_fit(load_config())
    print(f"[{source}] smart paste:", outcome)


def _on_smart_hotkey() -> None:
    _smart_paste("ctrl+alt+v")


def _on_ctrl_v() -> None:
    global _passthrough
    import keyboard  # type: ignore

    if _passthrough:
        return

    # Setting may have been turned off between registration and keypress.
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

    _smart_paste("ctrl+v")


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


def start_hotkeys(*, blocking: bool = True, on_status: Callable[[str], None] | None = None) -> None:
    """Register hotkeys. Safe to call once; later calls are no-ops."""
    global _started
    import keyboard  # type: ignore

    with _lock:
        if _started:
            if on_status:
                on_status("Hotkeys already running")
            return
        _started = True

    keyboard.add_hotkey("ctrl+alt+v", _on_smart_hotkey, suppress=False)
    _sync_ctrl_v_registration()
    threading.Thread(target=_watch_settings, name="wpif-hotkey-watch", daemon=True).start()

    msg = (
        "Hotkeys ready: Ctrl+Alt+V = smart paste; "
        "enable 'Hijack Ctrl+V' in the preview window to intercept image pastes in Word."
    )
    print(msg)
    if on_status:
        on_status(msg)

    if blocking:
        keyboard.wait()


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
    print("Open the preview window and enable 'Hijack Ctrl+V' if you want image-only Ctrl+V.")
    start_hotkeys(blocking=True)


if __name__ == "__main__":
    main()
