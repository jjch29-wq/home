import os
import glob
import olefile

target_folder = r"C:\Users\jjch2\Desktop\신고서_마스터템플릿\25.04.03_25.05.07_롯데건설(주) 롯데바이오로직스_변경신고(기간연장, 신고선원, 작업방법)\25.04.03_25.04.03_롯데건설(주) 롯데바이오로직스\롯데건설 신고서류_UG_최초"
output_file = r"c:\Users\jjch2\Desktop\PMI\hwp_all_text.txt"

hwp_files = glob.glob(os.path.join(target_folder, "*.hwp"))

with open(output_file, "w", encoding="utf-8") as out_f:
    for f in hwp_files:
        out_f.write(f"\n\n--- FILE: {os.path.basename(f)} ---\n")
        try:
            ole = olefile.OleFileIO(f)
            if ['PrvText'] in ole.listdir():
                stream = ole.openstream('PrvText')
                text = stream.read().decode('utf-16le')
                out_f.write(text)
            else:
                out_f.write("이 파일에는 텍스트 미리보기가 저장되지 않았습니다.\n")
            ole.close()
        except Exception as e:
            out_f.write(f"에러 발생: {e}\n")

print("SUCCESS")
