"""Date-local company notes, kept separate from common remarks and chemical makers."""


def memo_snapshot(memos, company=None, text=''):
    result = {str(name).strip(): value for name, value in (memos or {}).items()
              if str(name).strip() and isinstance(value, str) and value.strip()}
    if company:
        if text.strip():
            result[company] = text
        else:
            result.pop(company, None)
    return result


def memo_history_rows(history):
    rows = []
    for day, data in history.items():
        common = str(data.get('remarks') or '').strip()
        if common:
            rows.append((str(day), '', common))
        for company, text in memo_snapshot(data.get('company_memos', {})).items():
            rows.append((str(day), company, text))
    return sorted(rows, key=lambda row: row[0], reverse=True)


def report_memo_text(common, memos):
    parts = [f'[공통]\n{common}'] if common and common.strip() else []
    parts.extend(f'[{company}]\n{text}' for company, text in memo_snapshot(memos).items())
    text = '\n\n'.join(parts)
    if len(text) > 400 or len(text.splitlines()) > 12:
        return '전체 내용은 「업체별 메모」 시트 참조\n\n' + text[:280] + '\n…'
    return text
