"""Shared note contents for inquiry indicators and detail windows."""
import pandas as pd


def clean_text(value):
    return '' if pd.isna(value) else str(value).strip()


def memo_key(entry):
    date = pd.to_datetime(entry.get('Date', ''), errors='coerce')
    return (date.strftime('%Y-%m-%d') if not pd.isna(date) else '',
            clean_text(entry.get('Site', '')))


def daily_memo_lookup(frame):
    result = {}
    for _, entry in frame.iterrows():
        memo = clean_text(entry.get('DailyMemo', ''))
        key = memo_key(entry)
        if memo and key[0]:
            notes = result.setdefault(key, [])
            if memo not in notes:
                notes.append(memo)
    return {key: '\n\n'.join(notes) for key, notes in result.items()}


def record_notes(entry, lookup):
    return clean_text(entry.get('Note', '')), lookup.get(memo_key(entry), '')
