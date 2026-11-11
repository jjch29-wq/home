"""Review the participants of the current inspection without aggregating money."""


def review_participants(rows):
    participants = {}
    missing_times = []
    for row in rows:
        name = row['name'].strip()
        if not row['selected'] or not name:
            continue
        participants.setdefault(name, []).append(row['index'])
        if not row['time'].strip():
            missing_times.append(f"{row['index']:02d} {name}")

    warnings = []
    if missing_times:
        warnings.append('작업시간 미입력: ' + ', '.join(missing_times))
    duplicates = [f"{name} ({'·'.join(f'{i:02d}' for i in indices)}번)"
                  for name, indices in participants.items() if len(indices) > 1]
    if duplicates:
        warnings.append('중복 선택: ' + ', '.join(duplicates))
    return {
        'count': len(participants),
        'names': list(participants),
        'warnings': warnings,
    }
