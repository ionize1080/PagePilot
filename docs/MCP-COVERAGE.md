# MCP coverage · 0.4.0-rc.1

132 tools: all 120 upstream tools, 9 existing PagePilot tools, and 3 new object tools. Every v0.3.0 input schema and read-only annotation is preserved. See [object protocol](OBJECT-EDITING.md) and the full machine-readable `tool-inventory.json`.

New live coverage includes story/page scopes; section layout/columns and primary/first/even headers/footers; physical cells, split/merge, row/column editing and per-border formatting; fields, bookmarks, notes, TOC, content controls and custom styles; existing images, equations, text boxes, shapes, chart series and OLE properties. Object handles bind to the document and a normalized all-story Flat OPC hash. Mutations validate, back up, use a custom undo record and explicitly report recovery status.

Word version, document compatibility mode, protection and per-object COM support are all considered. Word 2010 does not expose threaded comments or repeating-section controls; a newer Word opening an old compatibility-mode file may still lack newer control types. Use the returned property schema and unavailableProperties rather than assuming every property exists.

Corrections include exact full-path document binding, cross-reference enum values, image option validation before insertion, and format/story-aware change detection. The new tools use the existing approval path; they cannot alter approval mode.

This is not full Word COM coverage or certification of every upstream behavior. OLE activation/internal editing, all chart axis/data-label properties, all OMath function nodes, native document comparison, task-level change attribution and SmartArt are not added. Word 2010 physical-machine validation remains pending. No numeric consistency business feature is introduced.
