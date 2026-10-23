# Third-party notices

PagePilot is distributed under GPL-3.0. This distribution includes the following software; its notices and licenses remain applicable.

* **Chatbox 1.20.3**, chatboxai/chatbox, GPL-3.0. Tag `v1.20.3`, commit `d3cd0ad1e6a1127bbbf3991bb5296c181f4cbdeb`. The official binary is extracted, never installed. Reproducible patches are in `chatbox/prepare.cjs`, `bootstrap.cjs`, and `preload.cjs`. `Chatbox/provenance.json` records before/after hashes. The release source archive includes the exact upstream Chatbox source archive.
* **word-mcp-live 1.6.21**, Yüce Karapazar and GongRzhe, MIT. Exact vendored source and license: `vendor/word-mcp-live`. Built upon GongRzhe/Office-Word-MCP-Server. Attribution does not imply endorsement.
* **Electron 44.7.0**, Electron contributors, MIT and Chromium third-party licenses. Preserve `Browser/LICENSE*` and `Browser/LICENSES.chromium.html`.
* **Python 3.12.10**, Python Software Foundation. Preserve `Python/LICENSE.txt`.
* Python dependencies are locked in `requirements-lock.txt`; distribution metadata and licenses remain in `Python/Lib/site-packages/*.dist-info`.
* Microsoft Office interop metadata is embedded from the Office 2010 PIAs and Extensibility reference assembly. Office itself and the .NET Framework are not bundled. Microsoft Word must be separately licensed.

Upstream downloads: https://github.com/chatboxai/chatbox/releases/tag/v1.20.3 , https://github.com/ykarapazar/word-mcp-live , https://www.python.org/downloads/release/python-31210/ , https://github.com/electron/electron/releases/tag/v44.7.0 .
