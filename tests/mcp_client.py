"""Small stdio-only MCP test driver. Session secrets stay in child environment."""
import json
import os
import queue
import subprocess
import threading
from pathlib import Path


class Client:
    def __init__(self, root, session):
        self.root = Path(root)
        env = dict(os.environ, PAGEPILOT_PIPE=session["pipe"], PAGEPILOT_TOKEN=session["token"], PYTHONIOENCODING="utf-8")
        self.proc = subprocess.Popen([str(self.root / "Python/python.exe"), str(self.root / "bridge/server.py")],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", env=env,
            creationflags=subprocess.CREATE_NO_WINDOW)
        self.queue = queue.Queue()
        self.errors = []
        threading.Thread(target=lambda: [self.queue.put(line) for line in self.proc.stdout], daemon=True).start()
        threading.Thread(target=lambda: [self.errors.append(line) for line in self.proc.stderr], daemon=True).start()
        self.sequence = 0
        self.request("initialize", {"protocolVersion": "2025-03-26", "capabilities": {}, "clientInfo": {"name": "PagePilot QA", "version": "0.2.0"}})
        self.proc.stdin.write(json.dumps({"jsonrpc":"2.0", "method":"notifications/initialized"})+"\n")
        self.proc.stdin.flush()

    def request(self, method, params):
        self.sequence += 1
        self.proc.stdin.write(json.dumps(dict(jsonrpc="2.0", id=self.sequence, method=method, params=params))+"\n")
        self.proc.stdin.flush()
        while True:
            try:
                response=json.loads(self.queue.get(timeout=180))
            except queue.Empty:
                raise RuntimeError("MCP timeout: " + "".join(self.errors)[-3000:])
            if response.get("id")==self.sequence:
                if "error" in response:
                    raise RuntimeError(str(response["error"]))
                return response["result"]

    def call(self, name, arguments=None):
        return self.request("tools/call",dict(name=name,arguments=arguments or {}))

    def close(self):
        self.proc.terminate()
        self.proc.wait(timeout=10)


def current_session():
    import win32com.client
    app=win32com.client.GetActiveObject("Word.Application")
    value=app.COMAddIns.Item("PagePilot.Connect").Object.GetMcpSession
    return json.loads(value() if callable(value) else value)
