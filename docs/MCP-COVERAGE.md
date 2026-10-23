# MCP coverage · v0.2.0

The complete machine-readable schema is `tool-inventory.json`: **120 upstream tools + 9 PagePilot tools = 129**, with no upstream name or parameter removed. This is interface coverage, not proof that every upstream operation has been executed on every Word version.

| Area | Exposed implementation | Boundary / added coverage |
|---|---|---|
| Document text, selection, search, replacement, insertion | upstream file and `word_live_*` tools | Open files use live COM. File tools require explicit paths; close files before file-based edits. |
| Fonts, paragraph spacing/alignment, styles, lists | upstream live formatting; `pagepilot_normalize_ascii`, `pagepilot_apply_hierarchy` | Native outline levels and optional multilevel numbering; ASCII normalization preserves Chinese runs. |
| Page/section layout, margins, headers, footers, page numbers | upstream live layout | Word/document state determines availability. |
| Tables, cells, borders, padding, dimensions, merging | upstream table tools; `pagepilot_normalize_tables` | Complex merged/nested table edge cases are not fully validated. |
| Comments, revisions, bookmarks, hyperlinks, images | upstream tools | Reply/resolve comments require Word 2013+, blocked on Word 2010 before approval. |
| Captions | `pagepilot_insert_caption` | Uses Word-native SEQ fields. |
| Footnotes, endnotes, TOC, fields | upstream file tools + `pagepilot_edit_objects` | Adds live insertion of footnotes/endnotes/TOC and field/TOC updating. |
| Content controls and object inspection | `pagepilot_inspect_objects`, `pagepilot_edit_objects` | Inspect fields/bookmarks/controls and counts; add rich-text controls and replace by inspected ID. |
| AI document structure | `pagepilot_context`, `pagepilot_set_navigation` | Read-only tree; explicit hierarchy tool is required to alter outline levels. |
| Version guidance | `pagepilot_capabilities` | Returns every tool name with available/conditional/limited/unavailable status, Word version, open/protected document state, and usage guidance. |
| File creation/conversion/protection/metadata | upstream tools | Bound-host PDF export avoids quitting the user's Word process. Signature metadata is not a certificate-backed signature. |

The AI receives instructions to read capabilities and context first. Tool availability is a conservative version/environment guide; it does not replace checking tool errors or document-specific constraints. No generic script execution escape hatch is exposed. All mutating MCP tools pass through host approval; read-only operations and direct user clicks in regex/navigation do not.

Word 2013 introduced comment replies and “done” status: [Microsoft's Word 2013 overview](https://www.microsoft.com/en-us/microsoft-365/blog/2012/07/30/introducing-word-2013/). Positioning uses [Word Window.ScrollIntoView](https://learn.microsoft.com/en-us/office/vba/api/word.window.scrollintoview), then adjusts within the actual document viewport. Page boundaries can limit the requested whitespace.

Target Win10 21H1 + Word 2010 has not been available for execution testing. The version matrix test simulates Word 14/16 metadata only and is not a substitute for that target test.
