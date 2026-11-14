"""Keep a selected material's identity separate from its editable display label."""
import pandas as pd


def _id(value):
    if value is None or pd.isna(value):
        return ''
    return str(value).strip().removesuffix('.0')


def _active(row):
    value = pd.to_numeric(row.get('Active', 1), errors='coerce')
    return pd.isna(value) or value > 0.5


def remember_selection(app, material_id):
    app._daily_material_selection = None
    if not _id(material_id):
        return
    display = app.cb_daily_material.get().strip()
    rows = app.materials_df[app.materials_df['MaterialID'].map(_id) == _id(material_id)]
    if len(rows) == 1 and _active(rows.iloc[0]) and app.get_material_display_name(rows.iloc[0]['MaterialID']) == display:
        app._daily_material_selection = {'id': _id(material_id), 'display': display}


def choose_daily_material(app, material_id):
    display = app.get_material_display_name(material_id)
    app.cb_daily_material.delete(0, 'end')
    app.cb_daily_material.insert(0, display)
    remember_selection(app, material_id)


def selected_id(app):
    selected = getattr(app, '_daily_material_selection', None)
    return selected['id'] if selected and selected['display'] == app.cb_daily_material.get().strip() else None


def resolve_daily_material(app, display=None):
    display = app.cb_daily_material.get().strip() if display is None else display.strip()
    if not display:
        return None
    selected = getattr(app, '_daily_material_selection', None)
    if selected and selected['display'] == display:
        rows = app.materials_df[app.materials_df['MaterialID'].map(_id) == selected['id']]
        if len(rows) != 1 or not _active(rows.iloc[0]) or app.get_material_display_name(rows.iloc[0]['MaterialID']) != display:
            raise ValueError('선택한 자재 정보가 변경되었습니다. 자재 검색에서 다시 선택하세요.')
        return rows.iloc[0]['MaterialID']
    app._daily_material_selection = None
    matches = [row['MaterialID'] for _, row in app.materials_df.iterrows()
               if _active(row) and app.get_material_display_name(row['MaterialID']) == display]
    if len(matches) > 1:
        raise ValueError('표시명이 같은 자재가 여러 개입니다. 자재 검색에서 ID를 확인해 선택하세요.')
    return matches[0] if matches else None
