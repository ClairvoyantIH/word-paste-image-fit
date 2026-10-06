from __future__ import annotations

"""Optional global hotkey helper.

MVP launches the preview window; hotkeys can be started with:
    python -m wpif.hotkeys
"""

from .config import load_config
from .word_paste import paste_and_fit, word_is_foreground


def main() -> None:
    from pynput import keyboard

    cfg = load_config()

    def on_smart_paste() -> None:
        # Always allow when Word is focused; otherwise still try active Word instance.
        if not word_is_foreground():
            # Soft gate: user asked for Word workflow; still attempt COM paste.
            pass
        outcome = paste_and_fit(load_config())
        print("smart paste:", outcome)

    # Default: Ctrl+Alt+V
    with keyboard.GlobalHotKeys({"<ctrl>+<alt>+v": on_smart_paste}) as h:
        print("Listening for Ctrl+Alt+V. Ctrl+C to exit.")
        print("hijack_ctrl_v setting is stored but not enabled in MVP hotkey helper:", cfg.hijack_ctrl_v)
        h.join()


if __name__ == "__main__":
    main()
