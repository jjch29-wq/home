import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox

MATERIAL_COST = {
    'RT (B필름: 3⅓"x17")': 8864,
    'RT (A필름: 3⅓"x12")': 8024,
    'RT (A/2필름: 3⅓"x6")': 6999,
    "UT": 1115,
    "PT": 3971
}

def update_ndt_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        records = data.get("records", [])
        updated_count = 0
        
        for r in records:
            if r["ndt_type"] == "RT":
                mat_type = str(r.get("material_type", "")).lower()
                if "b" in mat_type or "17" in mat_type:
                    mat_unit_cost = 8864
                elif "a/2" in mat_type or "6" in mat_type:
                    mat_unit_cost = 6999
                else:
                    mat_unit_cost = 8024
                    
                # 재료비는 실물량(qty) 기준
                new_mat_cost = int(r["qty"] * mat_unit_cost)
                
                # 기존 재료비와 다르면 업데이트 (소계도 차액만큼 더해줌)
                if r.get("mat_cost", 0) != new_mat_cost:
                    diff = new_mat_cost - r.get("mat_cost", 0)
                    r["mat_cost"] = new_mat_cost
                    r["subtotal"] = r.get("subtotal", 0) + diff
                    updated_count += 1
                    
        # 계약 총액 및 전회 총액 등은 앱에서 "프로젝트 총 계약수량 자동입력" 버튼으로 덮어씌워야 정확함.
        # 여기서는 레코드만 업데이트 후 저장
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        return updated_count
    except Exception as e:
        print(f"Error updating {filepath}: {e}")
        return -1

def main():
    root = tk.Tk()
    root.withdraw()
    
    # GUI를 최상단으로 끌어올림
    root.attributes('-topmost', True)
    
    filepaths = filedialog.askopenfilenames(
        title="단가를 업데이트할 NDT 저장 파일들을 모두 선택하세요 (.ndt)",
        filetypes=[("NDT Save Files", "*.ndt")]
    )
    
    if not filepaths:
        return
        
    total_files = 0
    total_records = 0
    
    for fp in filepaths:
        count = update_ndt_file(fp)
        if count > 0:
            total_files += 1
            total_records += count
            
    msg = (f"업데이트가 완료되었습니다!\n\n"
           f"선택한 파일 중 {total_files}개의 파일에서 총 {total_records}개의 과거 RT 검사 기록이 "
           f"동위원소 단가가 포함된 새로운 단가(8,864원 등)로 재계산되었습니다.\n\n"
           f"※ 주의: 메인 앱에서 이 파일들을 불러오신 후, 반드시 2번 탭 상단의 "
           f"'[프로젝트 총 계약수량 자동입력]' 버튼을 한 번 눌러서 계약 총액을 갱신해 주셔야 완벽하게 적용됩니다.")
           
    messagebox.showinfo("과거 기록 소급 업데이트 완료", msg)

if __name__ == "__main__":
    main()
