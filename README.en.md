# Word Paste Image Fit

Paste screenshots into Word without endless dragging.

An open-source Windows utility for **desktop Microsoft Word**: smart-fit pasted images to your style rules, with a **live print-layout style preview** so every parameter change shows what the next paste will look like.

[简体中文](./README.md) · [Roadmap](./docs/ROADMAP.md) · [Table-cell fitting](./docs/TABLE_CELL.md)

---

## Why this exists

Full-screen screenshots pasted into Word often arrive at raw pixel size: wider than the page, larger than a table cell, and painful to resize by hand.

Word’s built-in “In line with text” paste can shrink to the content width. This project goes further when you want:

- Width **and** height caps (so tall captures don’t eat the whole page)
- Rules like “92% of content width” or “fixed 14 cm”
- Fitting to a **table cell** instead of the full page
- Seeing the result **before** you paste

![Built-in sample screenshot](./assets/sample_screenshot.png)

## Features (MVP)

| Feature | What you get |
|---------|----------------|
| Live layout preview | Independent window; sliders update the page mock instantly |
| Built-in sample image | Tune styles without relying on the clipboard |
| Width + height limits | Avoid “fits width, blows past page height” |
| Named presets | Notes / Paper / Fixed 14cm (extendable) |
| Table-cell adaptive fit | When the caret is in a cell, size to that cell |
| zh / en UI | Switch the interface language |
| Smart paste | Paste into Word via COM and apply the active style |
| Non-invasive | **Does not hijack Ctrl+V by default** |

Settings live in `%AppData%\word-paste-image-fit\settings.json`.

## Quick start

Needs Windows, desktop Word, and Python 3.

```powershell
cd word-paste-image-fit
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\run_preview.ps1
```

Or launch **Word Paste Image Fit** from the Start Menu after install.

Suggested flow:

1. Drag max-width / max-height sliders and watch the preview.
2. Optionally enable “Preview: simulate table cell”.
3. Open Word and click **Smart paste into Word**
   (clipboard image if present; otherwise the built-in sample).
4. Optional hotkey helper: `$env:PYTHONPATH="src"; .\.venv\Scripts\python.exe -m wpif.hotkeys` (`Ctrl+Alt+V`).

## Table-cell adaptive fitting

If the caret is inside a table cell, fitting uses the cell box—not the full page content width. Details: [docs/TABLE_CELL.md](./docs/TABLE_CELL.md).

## Project layout

```
src/wpif/     Preview UI + fit math + Word paste
config/       Default style presets
assets/       Built-in sample screenshot
vba/          VBA mirror for a later .dotm path
scripts/      Install / launch helpers
docs/         Roadmap and concept notes
tests/        Geometry unit tests
```

## Roadmap

- **Phase A (now):** Python preview studio + COM smart paste
- **Phase B:** Better packaging, auto-update, `.dotm`
- **Phase C:** Native COM/VSTO add-in

Full checklist: [docs/ROADMAP.md](./docs/ROADMAP.md)

Non-goals for now: macOS Word, Word Online, WPS-first support.

## Contributing

Issues and PRs welcome—especially preview UX, paste edge cases, install polish, and a safe image-only Ctrl+V option.

## License

[MIT](./LICENSE)
