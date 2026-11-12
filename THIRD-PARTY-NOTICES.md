# Third-party notices

PagePilot is distributed under GPL-3.0. This distribution includes the following software; its notices and licenses remain applicable.

* **Chatbox 1.20.3**, chatboxai/chatbox, GPL-3.0. Tag `v1.20.3`, commit `d3cd0ad1e6a1127bbbf3991bb5296c181f4cbdeb`. The official binary is extracted, never installed. Reproducible patches are in `chatbox/prepare.cjs`, `bootstrap.cjs`, and `preload.cjs`. `Chatbox/provenance.json` records before/after hashes. The release source archive includes the exact upstream Chatbox source archive.
* **word-mcp-live 1.6.21**, Yüce Karapazar and GongRzhe, MIT. Exact vendored source and license: `vendor/word-mcp-live`. Built upon GongRzhe/Office-Word-MCP-Server. Attribution does not imply endorsement.
* **Electron 44.7.0**, Electron contributors, MIT and Chromium third-party licenses. Preserve `Browser/LICENSE*` and `Browser/LICENSES.chromium.html`.
* **Python 3.12.10**, Python Software Foundation. Preserve `Python/LICENSE.txt`.
* Python dependencies are locked in `requirements-lock.txt`; distribution metadata and licenses remain in `Python/Lib/site-packages/*.dist-info`.
* Microsoft Office interop metadata is embedded from the build-selected PIA and Extensibility reference assembly. Version 0.4.8 was compiled with the locally available Office 15 PIA; the base API target remains Word 14. Office itself and the .NET Framework are not bundled. Microsoft Word must be separately licensed.
* **pypdfium2 5.15.2**, BSD-3-Clause / Apache-2.0, with bundled **PDFium** and its third-party dependencies. Preserve the entire `Python/Lib/site-packages/pypdfium2-5.15.2.dist-info/licenses` tree, including `BUILD_LICENSES/pdfium.txt` and the dependency licenses. See [upstream distribution guidance](https://pypdfium2.readthedocs.io/en/stable/readme.html#licensing).
* **openpyxl 3.1.5**, **et_xmlfile 2.0.0**, **python-pptx 1.0.2**, **xlrd 2.0.2**, **XlsxWriter 3.3.1**, and **striprtf 0.0.35** provide local document extraction and working-paper writing. Their original license files are preserved in each distribution's `.dist-info` directory. These tools do not include or license Microsoft Office.

Upstream downloads: https://github.com/chatboxai/chatbox/releases/tag/v1.20.3 , https://github.com/ykarapazar/word-mcp-live , https://www.python.org/downloads/release/python-31210/ , https://github.com/electron/electron/releases/tag/v44.7.0 .

* **GridPilot** (ionize1080/GridPilot): the icon drawing approach, owned Electron window docking and composer approval interaction were adapted from the user-requested reference project. PagePilot supplies its own document/search/bookmark/gear artwork and Word integration.
