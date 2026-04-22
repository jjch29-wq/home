import pandas as pd
import re

def test_normalization():
    mat_id = 100.0
    str_mat_id_wrong = str(mat_id).replace('.0', '')
    str_mat_id_right = re.sub(r'\.0$', '', str(mat_id))
    
    print(f"Original: {mat_id}")
    print(f"Wrong (replace): '{str_mat_id_wrong}'")
    print(f"Right (regex): '{str_mat_id_right}'")

    mat_ids = [10.0, 20.0, 100.0, 101.0, 5.0]
    for mid in mat_ids:
        wrong = str(mid).replace('.0', '')
        right = re.sub(r'\.0$', '', str(mid))
        print(f"{mid} -> Wrong: '{wrong}', Right: '{right}'")

if __name__ == "__main__":
    test_normalization()
