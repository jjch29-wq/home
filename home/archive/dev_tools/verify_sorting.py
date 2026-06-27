
import re

def to_float(val):
    if val is None: return 0.0
    s = str(val).upper().replace("%", "").strip()
    if "<" in s or "ND" in s or s == "": return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def get_value(item, key):
    val = item.get(key, "")
    if key in ["Ni", "Cr", "Mo", "Mn"]:
        return to_float(val)
    if key == "No":
        try: 
            cleaned = re.sub(r'[^0-9.]', '', str(val))
            return float(cleaned) if cleaned else 0.0
        except: return str(val).lower()
    return str(val).lower()

def test_sorting():
    data = [
        {'No': '10', 'Dwg': 'ISO2', 'Ni': 10.0},
        {'No': '2', 'Dwg': 'ISO1', 'Ni': 8.0},
        {'No': '1', 'Dwg': 'ISO2', 'Ni': 12.0},
        {'No': '20', 'Dwg': 'ISO1', 'Ni': 9.0},
    ]

    # Test 1: Sort by No Ascending
    data.sort(key=lambda x: get_value(x, "No"), reverse=False)
    print(f"Sort by No Asc: {[x['No'] for x in data]}")
    assert [x['No'] for x in data] == ['1', '2', '10', '20']

    # Test 2: Sort by ISO Ascending (secondary No Ascending)
    data.sort(key=lambda x: (get_value(x, "Dwg"), get_value(x, "No")), reverse=False)
    print(f"Sort by ISO Asc: {[x['No'] for x in data]}")
    # ISO1: 2, 20
    # ISO2: 1, 10
    assert [x['No'] for x in data] == ['2', '20', '1', '10']

    # Test 3: Sort by ISO Descending (secondary No Descending)
    data.sort(key=lambda x: (get_value(x, "Dwg"), get_value(x, "No")), reverse=True)
    print(f"Sort by ISO Desc: {[x['No'] for x in data]}")
    # ISO2: 10, 1
    # ISO1: 20, 2
    assert [x['No'] for x in data] == ['10', '1', '20', '2']

    print("All sorting logic tests passed!")

if __name__ == "__main__":
    test_sorting()
