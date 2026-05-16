# Yuninput IME

Yuninput is a Chinese input method project for Windows 11 TSF (Text Services Framework).

This repository provides a buildable and extensible IME framework, including the core engine, candidate window, configuration tool, installer scripts, and EXE/MSI packaging flow.

## Status

- Stage: In active development
- Target platform: Windows 11 x64
- Stack: C++ TSF Text Service + PowerShell installer scripts + Windows config tool
- Default dictionary profile: zhengma-all
- Positioning: Engineering-grade IME framework, not a finalized commercial product yet

## Current Features

- C++ DLL project with CMake + MSVC
- TSF Text Service COM baseline
- DllRegisterServer / DllUnregisterServer
- Candidate window rendering and caret tracking
- Candidate selection with keys 1-9
- Candidate pinning: Ctrl+1-9
- Candidate blocking: Ctrl+Delete
- Paging: - / =, [ / ], , / ., PgUp / PgDn
- Tab / Shift+Tab candidate navigation
- Left Shift tap toggles Chinese/English mode
- F10 toggles full-shape / half-shape character mode
- F2 or APPS opens the status menu (F2 falls back to config launch when menu is unavailable)
- Space commit, Enter exact-priority commit, Esc clear, Backspace edit
- Continuous commit with remaining-code handling
- Wildcard input uses '0' in 4-code patterns (only position 2 and/or 4)
- Persistent frequency learning
- Restored code-frequency learning for compatibility-source single-character short codes (one/two-code)
- Persistent context association learning
- zhengma-large-pinyin profile supports pinyin single-character query with Zhengma code hints
- Config tool for user entries, auto phrases, and context-learning management
- Table-dict import support (Fcitx / IBus style)
- EXE and MSI package generation

## Build

Requirements:

- Windows 11
- Visual Studio 2022 with C++ desktop workload
- CMake 3.21+

Build commands:

```powershell
cmake -S . -B build -A x64
cmake --build build --config Release
```

## Install / Uninstall

Install (elevated PowerShell):

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/install_enable.ps1 -Build
```

Uninstall:

```powershell
./scripts/unregister_ime.ps1
```

## Dictionary and Learning

- Dictionary files are under data/
- User frequency and user dictionary are persisted
- Auto phrase learning is independent from manual user dictionary entries
- Session auto-phrase history is retained within a rolling 2000-Han window for cross-session continuity
- A phrase recalled from that window can be promoted into the user dictionary after one confirmed commit

## Licensing and Notices

- License: GPL-3.0-or-later
- See LICENSE and LICENSE.zh-CN.md
- Ownership and notices: NOTICE.md / NOTICE.zh-CN.md
- Third-party notices: THIRD_PARTY_NOTICES.md / THIRD_PARTY_NOTICES.zh-CN.md

## Chinese Documentation

For the Chinese main documentation, see README.md.
