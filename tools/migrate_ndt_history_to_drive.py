"""Initialize the configured Drive folder and migrate local issue history."""
import json
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / 'home/src/report_apps/ndt_report/src'))
import report_issue_history as history

if __name__ == '__main__':
    target = history.resolve_history_path()
    directory = target.parent
    if not directory.is_dir():
        raise OSError(f'Google Drive 데이터 폴더에 연결할 수 없습니다: {directory.parent}')
    directory.mkdir(exist_ok=True)
    records = history.load_history()
    print(f'발행 이력 {len(records)}건 확인: {target}')
