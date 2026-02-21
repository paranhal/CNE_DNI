from pathlib import Path
import json
import pandas as pd

cfg = Path('config/paths.local.json')
print('config exists:', cfg.exists())

if cfg.exists():
    c = json.loads(cfg.read_text(encoding='utf-8-sig'))
    print('RAW_ROOT:', c.get('RAW_ROOT'))
    print('OUT_ROOT:', c.get('OUT_ROOT'))
    print('LOG_ROOT:', c.get('LOG_ROOT'))

print('pandas:', pd.__version__)
print('OK')
