"""차량과 검사방법 연결 규칙."""


def normalized_method(value):
    return str(value or '').strip().upper()


def vehicle_applies_to_method(vehicle_data, current_method):
    vehicle_method = normalized_method(vehicle_data.get('inspection_method', ''))
    return not vehicle_method or vehicle_method == normalized_method(current_method)


def validate_vehicle_methods(vehicle_rows, item_methods):
    methods = {normalized_method(value) for value in item_methods if normalized_method(value)}
    if len(methods) <= 1:
        return ''
    for vehicle in vehicle_rows:
        if not str(vehicle.get('vehicle_info', '')).strip():
            continue
        method = normalized_method(vehicle.get('inspection_method', ''))
        if not method:
            return '다중 검사 저장 시 차량별 검사방법을 선택해주세요.'
        if method not in methods:
            return f"차량 검사방법 {method}이(가) 다중 품목 검사방법에 없습니다."
    return ''


def vehicle_checks(vehicle_data):
    reserved = {'vehicle_info', 'mileage', 'remarks', '_raw_mileage', 'photo_path'}
    return '|'.join(
        f'{key}:{value}' for key, value in vehicle_data.items()
        if key not in reserved and value)


def update_vehicle_in_record(record, vehicle_data):
    """저장 행에 포함된 동일 차량의 점검 정보를 제자리에서 수정한다."""
    target = ''.join(str(vehicle_data.get('vehicle_info', '')).split()).upper()
    numbers = [part.strip() for part in str(record.get('차량번호', '') or '').split('||')]
    try:
        index = next(i for i, number in enumerate(numbers)
                     if ''.join(number.split()).upper() == target and target)
    except StopIteration:
        return False

    fields = {
        '주행거리': str(vehicle_data.get('_raw_mileage', '')).strip(),
        '차량점검': vehicle_checks(vehicle_data),
        '차량비고': str(vehicle_data.get('remarks', '')).strip(),
    }
    for column, value in fields.items():
        parts = [part.strip() for part in str(record.get(column, '') or '').split('||')]
        parts.extend([''] * (len(numbers) - len(parts)))
        parts[index] = value
        record[column] = ' || '.join(parts)
    return True
