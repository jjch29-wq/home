"""Local development isolation; production worktrees keep their existing paths."""
from pathlib import Path


def is_development():
    root = Path(__file__).resolve().parents[5]
    if (root / '.central-development-mode').exists():
        return True
    git_path = root / '.git'
    if git_path.is_file():
        pointer = git_path.read_text(encoding='utf-8').strip()
        if not pointer.startswith('gitdir: '):
            return False
        git_path = (root / pointer[8:]).resolve()
    head = git_path / 'HEAD'
    return head.exists() and head.read_text(encoding='utf-8').strip() == 'ref: refs/heads/중앙지사-개발'


def disabled_sync_result():
    return {'status': 'success', 'added_count': 0, 'pulled_count': 0,
            'pushed_count': 0, 'disabled': True,
            'message': '개발용: 운영 클라우드 동기화가 차단되어 있습니다.'}
