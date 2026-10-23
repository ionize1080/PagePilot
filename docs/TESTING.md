# PagePilot 0.1.0 validation

Date: 2026-10-23 (Asia/Singapore).

## Actual environment

* Windows build 26300, Word 16.0.20430.20180 x64.
* .NET Framework 4.8, Office 2010 interop metadata, Python 3.12.10 embedded x64, Electron 44.7.0.
* Chatbox 1.20.3 official installer/archive hashes verified before extraction; no installer run.
* Target Windows 10 21H1 + Word 2010 x86/x64: **not yet available for a real execution test**. The user confirmed on 2026-10-23 that no target test environment is currently available. Release notes explicitly retain this unverified compatibility limitation.

## Passed checks

* Native add-in activated on Word startup; native task pane and ribbon displayed.
* Separate WPF regex and navigation panes visually verified at 200% DPI, simultaneously open with independent close controls.
* Actual WPF regex UI previews two capture-group matches, obtains approval, and writes [2024]/[2025] to the fixture; COM readback confirms exact content and unchanged outline level.
* Electron settings reads and saves the session approval flag; its provider button opens the original Chatbox model settings.
* Closing the last document and opening a new one with the same Word window handle rebinds the pane, clears old regex previews, and selects the new heading; outline remains unchanged.
* Native caption extension inserts an actual Word SEQ field, verified through COM after stdio invocation.
* 11 regex/heading engine assertions: Chinese headings, decimal levels, body exclusion, offsets, named capture replacement, zero-width matches, Word paragraph anchors, paragraph/cell protection, regex timeout.
* stdio initialize / tools/list discovers all 120 upstream tools, plus PagePilot extensions; no filter drops upstream tools.
* Authenticated named-pipe context reads and actual `word_live_get_text` return Word document contents.
* Manual denial of `word_live_insert_text` returns an error and leaves the sentinel absent from the real document.
* Real DeepSeek API test with an externally supplied key: reads synthetic document, diagnoses 100→120 as 20% instead of 10%, executes native normalization on 11 runs, generates four navigation entries, and retains all original outline levels. COM readback confirms Times New Roman 10.5 pt and unchanged source wording.
* Actual Chatbox UI → DeepSeek → MCP → Word test successfully reads the document, identifies the arithmetic error and a 120 vs 125 table/body discrepancy, and returns evidence with uncertainty about whether the table uses the same reporting basis.
* Automatic approval regression exercises info, find, font formatting, paragraph spacing, page margins, table creation/normalization/formatting/cell text, comments, header/footer, page numbers, multilevel hierarchy, tracked insertion, revision listing, paragraph inspection, layout diagnostics, PDF export and save. Assertions inspect actual Word state, not just response messages. PDF output is nonempty and Word remains running with its original document.
* Upstream test suite contains one PDF conversion test; it passed. Its helper's application-wide Quit side effect was identified and replaced in PagePilot's gateway by a bound-host PDF export implementation.
* Observed PagePilot/Word/Python/Electron processes expose no listening TCP sockets during the tested session. PagePilot MCP is stdio only; internal communication is a named pipe.

## Reproducible entry points

* `tests/CoreTests.cs`: compile against PagePilot.dll and run with .NET Framework.
* `tests/live_smoke.py`: actual named pipe + stdio discovery/context; `--deny` waits for UI rejection.
* `tests/deepseek_live.py <key-file>`: synthetic-document real API loop; reads the secret in memory, never prints it or writes it into the report.
* `tests/operations.py`: real Word regression matrix; requires the disposable `PagePilot-QA.docx` and the Electron settings window's automatic approval checkbox.

## Remaining acceptance work

* Actual Win10 21H1 + Word 2010 x86 and x64 startup, pane, regex, stdio and editing tests.
* Per-tool destructive edge cases across every one of the 120 upstream tools. Tool/schema coverage is complete; per-tool behavioral testing is not.
* Large documents, multi-window close/reopen lifecycle, merged/nested tables and 100%/150%/mixed-monitor DPI regression.
* Word 2010 does not expose all modern comment-threading APIs. Those tools remain listed and must report unsupported interfaces; they are not silently presented as fully compatible.
* Upstream signature metadata is not a certificate-backed digital signature.
