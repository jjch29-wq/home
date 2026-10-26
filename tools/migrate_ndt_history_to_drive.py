"""Initialize the configured Drive folder and migrate local issue history."""
import json
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / 'home/src/report_apps/ndt_report/src'))
import report_issue_history as history

if __name__ == '__main__':
    config = json.loads(history.SYNC_CONFIG.read_text(encoding='utf-8'))
    directory = Path(config['data_dir'])
    if not directory.is_absolute() or not directory.parent.is_dir():
        raise OSError(f'Google Drive 데이터 폴더에 연결할 수 없습니다: {directory.parent}')
    directory.mkdir(exist_ok=True)
    records = history.load_history()
    print(f'발행 이력 {len(records)}건 확인: {directory / "report_issue_history.json"}')
