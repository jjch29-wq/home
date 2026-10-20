"""중앙지사 제출용 속성정보 생성기 실행."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / 'home' / 'src'))
from site_apps.central.src.attribute_export import main

if __name__ == '__main__':
    main()
