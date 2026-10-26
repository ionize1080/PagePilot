"""Operational metadata only; never serialize tool arguments or returned content."""
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
from datetime import datetime, timezone

_logger = None
def event(name, detail='', level='info'):
    global _logger
    if level == 'debug' and os.environ.get('PAGEPILOT_LOG_LEVEL') == 'info':
        return
    try:
        folder = os.environ.get('PAGEPILOT_LOGS')
        if not folder:
            return
        if _logger is None:
            Path(folder).mkdir(parents=True, exist_ok=True)
            _logger = logging.getLogger('pagepilot.operational')
            _logger.propagate = False
            _logger.setLevel(logging.INFO)
            handler = RotatingFileHandler(Path(folder) / f'bridge-{os.getpid()}.jsonl', maxBytes=4*1024*1024, backupCount=3, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(message)s'))
            _logger.addHandler(handler)
        _logger.info(json.dumps(dict(utc=datetime.now(timezone.utc).isoformat(), pid=os.getpid(), component='bridge', level=level, eventName=name, detail=str(detail)[:12000]), ensure_ascii=False))
    except Exception:
        pass
