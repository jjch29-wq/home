# -*- coding: utf-8 -*-
import pandas as pd
df = pd.read_excel('home/data/Material_Inventory.xlsx', sheet_name='현장사용내역')
print(df.tail(10))
