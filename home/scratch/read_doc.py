import sys, os, win32com.client
filepath = r'C:\Users\-\OneDrive\바탕 화면\KS B 0845(2026ed)Rev.0.doc'
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(filepath, ReadOnly=True, Visible=False)
text = doc.Content.Text
doc.Close(False)
word.Quit()
print(text[:2000])
