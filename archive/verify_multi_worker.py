
import pandas as pd
import datetime

# Mock data structure
columns = ['Date', 'Site', 'Material ID', 'Usage', 'Note', 'Entry Time', 
           'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
           'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사방법', '검사량',
           '단가', '출장비', '일식', '검사비', 'User', 'User2', 'User3', 'User4', 'User5', 'User6']

df = pd.DataFrame(columns=columns)

# Mock save logic
def mock_add_entry(df, users):
    new_entry = {
        'Date': datetime.datetime.now(),
        'Site': 'Test Site',
        'Material ID': 1,
        'User': users[0],
        'User2': users[1],
        'User3': users[2],
        'User4': users[3],
        'User5': users[4],
        'User6': users[5]
    }
    return pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

# Test 1: Save 6 workers
test_users = ["User A", "User B", "User C", "User D", "User E", "User F"]
df = mock_add_entry(df, test_users)

assert len(df) == 1
assert df.iloc[0]['User'] == "User A"
assert df.iloc[0]['User6'] == "User F"
print("Dataframe save verification SUCCESS")

# Test 2: Partial workers
test_users_partial = ["User G", "User H", "", "", "", ""]
df = mock_add_entry(df, test_users_partial)
assert df.iloc[1]['User2'] == "User H"
assert df.iloc[1]['User3'] == ""
print("Partial worker save verification SUCCESS")

# Test 3: Backward compatibility (simulate old file load)
old_df = pd.DataFrame(columns=['Date', 'Site', 'User'])
old_df.loc[0] = [datetime.datetime.now(), 'Old Site', 'Old User']

string_columns = ['Site', 'Note', 'User', 'User2', 'User3', 'User4', 'User5', 'User6', '장비명', '검사방법']
for col in string_columns:
    if col not in old_df.columns:
        old_df[col] = ''

assert 'User6' in old_df.columns
assert old_df.iloc[0]['User6'] == ""
print("Backward compatibility logic verification SUCCESS")
