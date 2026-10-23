# PagePilot 0.2.0 validation

Date: 2026-10-23. Actual host: Windows build 26300, Word 16.0.20430.20180 x64, .NET Framework 4.8, Office 2010 PIA; Chatbox 1.20.3 / Electron 44.7.0 / embedded Python 3.12.10.

**Win10 21H1 + Word 2010 x86/x64 remains untested. No target VM/machine is available.** Simulated Word 14 metadata tests do not establish target compatibility.

## Verified in this update

- Core engine: 11 assertions passed (headings, paragraph anchors, regex groups/offsets, boundaries, timeout).
- Capability matrix: Word 14 and 16; no document; protected document. All 120 upstream tools retained; 129 total schemas. Word 14 comment reply/resolve unavailable, signature metadata limited.
- Actual stdio discovery, host context and live Word text returned; new object tools verified by COM readback: footnote, endnote, rich-text content control insertion/replacement, TOC insertion/update, field update. Invalid operation rejected.
- Real DeepSeek API: capability/context inspection, numerical audit (100 to 120 is 20%, not 10%), Times New Roman 10.5 normalization and top-level-only navigation. API key read only from the supplied external file, not logged or packaged.
- Native Chatbox UI: exact requested audit prompt sent and capability/context calls returned; real API response identified 20%. Native navigation textarea requirements transferred to the actual React composer; exact personalized prompt sent, AI called navigation tool. COM confirmed unchanged document text/outline levels.
- Native AI window: six dock/undock transitions, retained Chinese draft and editable input. Docked window controls hidden; top/composer mode buttons present. No SetParent/child-window embedding.
- Automatic approval: cancel stays manual, explicit confirmation enables red lightning, disable restores manual. Default confirmation action is keep manual. Fixed confirmation timeout, busy-state and foreground visibility problems found in QA.
- Independent settings: all three palette selections and 72 px reading whitespace persisted; standalone settings process had only its settings window. Added actual four-icon previews. WPF accent and gear follow palette.
- Navigation: hierarchy displayed as a collapsible tree; single click collapses selection and positions heading near top, body visible below; repeated selected-heading click works. Page boundaries limit exact pixel positioning.
- Actual WPF regex UI replaced exactly 18 capture-group matches without approval. Manual outline button directly applied 1/2/1 levels; backups retained.
- Re-executed 20-operation stdio regression: font, spacing, page margins, table creation/normalization/cell edits, comments, headers/footers, page numbers, hierarchy, revisions, paragraph inspection, diagnostics, PDF and save; actual Word state asserted.
- Switching to a new document resets automatic approval; returning does not re-enable it. Manual denial of a stdio insert leaves DENIED_SENTINEL absent in actual Word.
- Final normal-runtime process snapshot: 7 related Word/Electron/Python processes, 0 listening TCP sockets. Restarting Word and opening settings displayed only the settings Electron window, no Chatbox.

## Reproduce

`tests/CoreTests.cs`, `tests/capabilities.py`, `tests/live_smoke.py`, `tests/objects_live.py`, `tests/deepseek_live.py <key-file>`, `tests/operations.py`.

`tests/electron-ui.cjs` and `tests/chatbox_api.cjs` are test-only staging injections loaded by a temporary bootstrap. They use Electron's own DOM/IPC and native input without a debugger server. Restore the unmodified repository bootstrap before packaging. The synthetic document is PagePilot-v020-QA.docx. Set PAGEPILOT_QA_OUTPUT to an isolated test folder; no real user document is used.

## Limits

129 advertised interfaces are not 129 individually exhaustive behavior tests. Large documents, mixed-monitor DPI, multiple Word processes, complex nested/merged tables, and all upstream destructive edge cases remain further acceptance work. The local build succeeded; NuGet vulnerability-feed retrieval was unavailable (NU1900), not counted as a passed security audit.

Prior-release tests are retained in TESTING-v0.1.0.md as historical evidence, not silently counted as new executions.
