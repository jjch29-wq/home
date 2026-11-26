"""Bootstrap local runtime configuration without shipping user settings."""
import json
import os
from pathlib import Path


def ensure_initial_config(path):
    """Create an empty valid JSON config once; never replace an existing file."""
    target = Path(path)
    if target.exists():
        return False
    target.parent.mkdir(parents=True, exist_ok=True)
    created = False
    try:
        with target.open('x', encoding='utf-8') as stream:
            created = True
            json.dump({}, stream, ensure_ascii=False, indent=2)
            stream.write('\n')
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError:
        return False
    except Exception:
        if created:
            try:
                target.unlink()
            except OSError:
                pass
        raise
    return True
