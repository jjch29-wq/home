import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc_path = r"C:\Users\-\OneDrive\문서\카카오톡 받은 파일\UT(23 Ed) B31.1 영한.doc"
out_path = r"C:\Users\-\PMI\home\UT(23 Ed) B31.1 영한_강조표시.doc"

try:
    doc = word.Documents.Open(doc_path)
    
    # 횡방향 관련 핵심 키워드 지정
    phrases = [
        "parallel and transverse directions", 
        "평행 및 횡 방향",
        "scanning parallel to the weld axis"
    ]
    
    for phrase in phrases:
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Text = phrase
        rng.Find.Forward = True
        rng.Find.Wrap = 0 # wdFindStop
        
        while rng.Find.Execute():
            rng.HighlightColorIndex = 7 # wdYellow
            rng.Collapse(Direction=0) # wdCollapseEnd
            
    doc.SaveAs(out_path)
    doc.Close(False)
    print(f"Saved to {out_path}")
    os.startfile(out_path)
except Exception as e:
    print("Error:", e)
finally:
    word.Quit()
