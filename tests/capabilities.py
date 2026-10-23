"""Version guidance is simulated; this is not a Word 2010 execution test."""
import asyncio
import importlib.util
from pathlib import Path
spec=importlib.util.spec_from_file_location('bridge',Path(__file__).resolve().parents[1]/'bridge/server.py')
bridge=importlib.util.module_from_spec(spec);spec.loader.exec_module(bridge)
async def main():
    for major,opened,protected in [(14,True,False),(16,True,False),(16,False,False),(16,True,True)]:
        environment=dict(wordVersion=str(major),majorVersion=major,documentOpen=opened,protectedDocument=protected)
        bridge.host=lambda action,**kw:environment
        result=await bridge.pagepilot_capabilities()
        catalog={t['name']:t for t in result['tools']}
        assert len(catalog)==129
        assert catalog['add_digital_signature']['status']=='limited'
        if major==14:
            assert all(catalog[n]['status']=='unavailable' for n in ['word_live_reply_to_comment','word_live_resolve_comment'])
        elif opened and not protected:
            assert catalog['word_live_reply_to_comment']['status']=='available'
        if not opened or protected:
            assert catalog['pagepilot_edit_objects']['status']=='conditional'
        print('PASS capability matrix',major,opened,protected)
    upstream={t['name'] for t in result['tools'] if not t['name'].startswith('pagepilot_')}
    assert len(upstream)==120
    print('PASS all 120 upstream tools retained')
asyncio.run(main())
