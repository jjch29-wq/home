"""Preserve registered vehicle descriptions and non-vehicle travel choices."""


def vehicle_label(value):
    value = str(value).strip()
    if '(' in value or value == '자가용/대중교통':
        return value
    kind = '탑차' if value in {'81두1580', '89보4028', '90너4889'} else '스타렉스'
    return f'{value} ({kind})' if value else ''


def restore_vehicle_catalog(config):
    defaults = ['84저1259', '95가0200', '91주8839', '90서8616',
                '76마3422', '71고4405', '89보4028', '81두1580',
                '81루5100', '95가0175', '81도5958', '90너4889']
    return sorted({vehicle_label(v) for v in config.get('vehicles', defaults) if str(v).strip()})
