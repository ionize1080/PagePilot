import json
import sys
from pathlib import Path
from mcp_client import Client, current_session

root=Path(__file__).resolve().parents[1]/"dist"
c=Client(root,current_session())
try:
    tools=c.request("tools/list",{})["tools"]
    print("Discovered tools:",len(tools),flush=True)
    assert len(tools)==129
    context=c.call("pagepilot_context")
    assert not context.get("isError"),context
    print("Context:",str(context)[:500],flush=True)
    result=c.call("word_live_get_text",{})
    assert not result.get("isError"),result
    print("Live read:",str(result)[:600],flush=True)
    if "--deny" in sys.argv:
        print("Waiting for manual rejection of test mutation",flush=True)
        try:
            result=c.call("word_live_insert_text",{"text":"DENIED_SENTINEL","position":"end"})
            assert result.get("isError"),result
        except RuntimeError as e:
            assert "User denied" in str(e),str(e)
        print("Manual rejection enforced",flush=True)
finally:
    c.close()
