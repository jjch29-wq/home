import xlsxwriter
import os
import glob
from PIL import Image
import sys

# [수정] 스크립트가 실행되는 위치를 기준으로 경로 설정 (경로 오류 방지)
base_path = os.path.dirname(os.path.abspath(__file__))
image_folder = os.path.join(base_path, 'images')
output_filename = os.path.join(base_path, 'NDT_Photo_Log_Final_v48.xlsx')
logo_filename = 'logo.png'

# 폴더가 없는 경우 대비
if not os.path.exists(image_folder):
    print(f"오류: '{image_folder}' 폴더를 찾을 수 없습니다. 폴더를 생성해 주세요.")
    sys.exit()

# 모든 사진 확장자 (대소문자 포함) 검색
all_files = sorted(glob.glob(os.path.join(image_folder, '*.[jJ][pP][gG]')) + 
                   glob.glob(os.path.join(image_folder, '*.[pP][nN][gG]')) +
                   glob.glob(os.path.join(image_folder, '*.[jJ][pP][eE][gG]')) +
                   glob.glob(os.path.join(image_folder, '*.[bB][mM][pP]')))

# 로고 파일은 사진 목록에서 제외
image_files = [f for f in all_files if os.path.splitext(os.path.basename(f))[0].lower() != 'logo']

print(f"작업 경로: {base_path}")
print(f"발견된 사진 개수: {len(image_files)}장")

if len(image_files) == 0:
    print("사진이 없습니다. 'images' 폴더에 사진을 넣었는지 확인해 주세요.")
    sys.exit()

# --- 이후 엑셀 생성 로직은 동일 ---
workbook = xlsxwriter.Workbook(output_filename)
worksheet = workbook.add_worksheet()

worksheet.set_paper(9) # A4
worksheet.set_portrait()
worksheet.set_margins(left=0.7, right=0.01, top=0.5, bottom=0.4)
worksheet.fit_to_pages(1, 0)
worksheet.repeat_rows(0, 4) 
worksheet.set_footer('&C&P / &N')

# 서식
title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'shrink': True})
company_format = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True})
center_border = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
desc_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_size': 10, 'shrink': True, 'text_wrap': False, 'indent': 1})

worksheet.set_column('A:B', 40)

# 헤더
worksheet.merge_range('A1:B1', "REPORT OF PHASED ARRAY UT EXAMINATION (위 상 배 열 초 음 파 탐 상 검 사 보 고 서)", title_format)
company_info_text = "서   울   檢   査   株   式   會   社\nSEOUL INSPECTION & TESTING Co., Ltd.\n서울특별시 서초구 바우뫼로 41길 54\nTEL : (02) 552-1112   FAX : (02) 2058-0720"
worksheet.merge_range('A2:A4', company_info_text, company_format)

# 로고
logo_path = os.path.join(base_path, logo_filename)
if not os.path.exists(logo_path):
    logo_path = os.path.join(image_folder, logo_filename)

if os.path.exists(logo_path):
    try:
        for r in range(1, 4): worksheet.set_row(r, 15)
        with Image.open(logo_path) as img:
            w, h = img.size
            scale = min(285/w, 45/h) * 0.6
            worksheet.insert_image('A2', logo_path, {'x_scale': scale, 'y_scale': scale, 'x_offset': 15, 'y_offset': 45 - (h * scale) - 2, 'object_position': 1})
    except: pass

worksheet.write('B2', "발주처: 서울에너지공사", center_border)
worksheet.write('B3', "REPORT NO: SIT/GI-SE-PAUT-TNTFJPWJ001", center_border)
worksheet.write('B4', "검사일자: 2024년 10월 24일", center_border)
worksheet.merge_range('A5:B5', "PHOTO LOG (사진 대장)", bold_format)

# 사진 삽입
row = 5
col = 0
CELL_WIDTH_PX = 286 
CELL_ROW_HEIGHT = 150
CELL_HEIGHT_PX = 200 
DESC_ROW_HEIGHT = 28.93

for image_path in image_files:
    worksheet.set_row(row, CELL_ROW_HEIGHT)
    try:
        with Image.open(image_path) as img:
            img_w, img_h = img.size
            x_scale = CELL_WIDTH_PX / img_w
            y_scale = CELL_HEIGHT_PX / img_h
            worksheet.insert_image(row, col, image_path, {'x_scale': x_scale, 'y_scale': y_scale, 'x_offset': 0, 'y_offset': 0, 'object_position': 1})
    except Exception as e:
        print(f"이미지 오류({os.path.basename(image_path)}): {e}")

    name_only = os.path.splitext(os.path.basename(image_path))[0]
    worksheet.set_row(row + 1, DESC_ROW_HEIGHT)
    worksheet.write(row + 1, col, f"설명: {name_only}", desc_format)
    
    if col == 0: col = 1
    else: col = 0; row += 2 

if col == 1:
    worksheet.write(row, 1, "", center_border)
    worksheet.write(row+1, 1, "설명:", desc_format)

workbook.close()
print(f"'{output_filename}' 파일이 완성되었습니다.")