# PagePilot 0.3.0 validation — 2026-10-23

Test host: physical Windows build 26300, Word 16 x64, Office14 PIA, .NET Framework 4.8.
The user explicitly excluded the VM from this test run. Word 2010 x86/x64 has not been executed; fallback code and Office14 compilation are not target-machine certification.

## Executed

- Release build passed. NU1900: vulnerability feed was unavailable; this is not a successful dependency vulnerability audit.
- Core engine: 17 assertions, including 100,001 matches, regex timeout, boundaries, hierarchy false positives and trailing dot depth.
- Actual DeepSeek model from the existing Chatbox configuration: inline navigation produced exactly the requested first-level headings at paragraph indices 1 and 5. No additional assistant window was launched. Cancel completed without applying results.
- Document XML equality passed with automatic proofing excluded from the fixture (Word independently splits runs during grammar checking).
- Actual Word regex: debounced search, single-click selection, capture/replacement, full batch undo, hyperlink-aware position mapping.
- Complex Word fixture: nested TOC/PAGEREF fields, footnotes, endnotes and content controls. Seven matches map to exact visible character ranges without field delimiters. Visible geometry and unchanged text verified.
- Actual on-screen review: yellow highlights on QA, amounts, dates and page-number results. Overlay does not alter formatting, save state or undo history.
- Chatbox UI: six dock/window transitions retained editable Chinese drafts; top-only visible mode switch; aligned input toolbar; closing approval confirmation preserved manual mode; confirmation enabled red state; disabling restored manual.
- Real Chatbox API numerical review returned 20% for 100 to 120. API keys remained in the existing user profile and were not logged or packaged.
- 20-operation actual stdio MCP regression: formatting, spacing, page setup, tables/cells, comments, headers/footers, page numbers, hierarchy, revisions, diagnostics, PDF and save. COM readback assertions passed.
- Object regression: footnote/endnote, content control insertion and update, TOC and field update, invalid operation rejection.
- Additional regression: context paging, bound screenshot, ASCII NameAscii assignment preserving 13 pt size, numbering-template reuse.
- Final approval regression: closing confirmation and reentrant requests stayed manual; explicit enable succeeded; switching documents reset approval permanently for that session.
- Capability matrix: 129 total tools including all 120 upstream tools, simulated Word 14/16 and protected/no-document states.
- JavaScript syntax, Python compilation and git whitespace checks.

## Reproduce

Build with build.ps1, then build tests/PhysicalIntegration.csproj. Copy the freshly built PagePilot.dll beside the test executable to avoid same-version assembly-cache reuse. The test uses the existing model configuration and only its new synthetic document. tests/ComplexSearch.csproj takes a disposable DOCX containing fields and controls. tests/operations.py and tests/objects_live.py require the named synthetic QA fixture and explicit automatic approval.

The Electron tests are staging-only injections; the release bootstrap must equal chatbox/bootstrap.cjs. No debugger port or test injection is shipped.

## Practical limits

- The regex engine remains .NET with its existing capture/replacement semantics and a 500 ms per-match timeout. A cancellation token stops superseded work; one million results is the memory protection limit.
- Highlights cover the active visible main-story viewport, with a 600-rectangle / 35 ms paint-work budget. Complex wrapping and unusual Word views depend on Word's GetPoint geometry.
- Complex nested fields may require character-by-character COM mapping. The 100,001-match microbenchmark measures matching only, not Word extraction or field mapping.
- Idle same-length document changes are detected approximately every 5.4 seconds; length changes are detected sooner. Every navigation/replacement verifies current content.
- Inline AI uses the selected HTTPS OpenAI-compatible chat-completions provider. Other provider protocols remain available through the full assistant. Batches preserve original paragraph indices; cross-batch hierarchy consistency remains model-dependent.
- Persistent scroll/highlight overlays are not written into the document. The overlay is hidden when Word loses foreground focus or the pane is closed.
