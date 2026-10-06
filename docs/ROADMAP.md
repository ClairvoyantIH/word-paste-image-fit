# Roadmap

## Phase A — MVP (current)

- [x] Independent settings + live page preview window
- [x] Built-in sample screenshot for instant feedback
- [x] Width + height constraints
- [x] Named style presets
- [x] zh / en UI language toggle
- [x] Smart paste into desktop Word via COM
- [x] Table-cell-aware fitting flag
- [x] Config persisted under `%AppData%\word-paste-image-fit\`
- [x] VBA mirror module for later `.dotm` packaging
- [ ] Optional global hotkey helper polish (`Ctrl+Alt+V`)
- [ ] Optional Ctrl+V hijack (images only)

## Phase B — Installer & packaging

- [x] One-click `install.ps1` Start Menu shortcut
- [ ] Signed release binaries / PyInstaller freeze
- [ ] Auto-update check (GitHub Releases)
- [ ] `.dotm` export path for macro-only environments

## Phase C — Native add-in

- [ ] COM / VSTO add-in reading the same JSON settings
- [ ] Task-pane optional; keep independent preview as advanced studio
- [ ] Per-document style override

## Non-goals (for now)

- macOS Word
- Word Online / Office.js clipboard hijack
- WPS-first support (may work via COM incidentally, not guaranteed)
