from PIL import Image, ImageFilter
import os

# 스크립트가 위치한 폴더 내의 signs 폴더를 지정합니다.
script_dir = os.path.dirname(os.path.abspath(__file__))
signs_dir = os.path.join(script_dir, 'signs')

def process_signature(filename, left_pad, out_filename, make_thick=False):
    """
    도장 이미지에 패딩(여백)을 추가하여 위치를 이동시키고,
    필요한 경우 흐릿한 서명 선을 진하고 굵게 만듭니다.
    """
    path = os.path.join(signs_dir, filename)
    if not os.path.exists(path):
        print(f"파일을 찾을 수 없습니다: {path}")
        return
        
    img = Image.open(path).convert('RGBA')
    
    # --- 선 굵기 및 진하기 조절 로직 ---
    if make_thick:
        data = img.getdata()
        new_data = []
        for item in data:
            # item[3]은 알파(투명도) 값입니다. 
            # 약간이라도 흔적이 있는 픽셀(>30)을 잡아내서 완전 검은색으로 진하게 칠합니다.
            if item[3] > 30: 
                new_alpha = min(255, int(item[3] * 5)) # 투명도를 확 끌어올림
                new_data.append((0, 0, 0, new_alpha)) # R, G, B를 0(검은색)으로 통일
            else:
                new_data.append((255, 255, 255, 0)) # 완전 투명
        img.putdata(new_data)
        
        # 선 두께를 굵게 만들기 위해 Alpha 채널(투명도 레이어)을 팽창시킵니다.
        r, g, b, a = img.split()
        a = a.filter(ImageFilter.MaxFilter(3)) # 3x3 필터를 써서 사방으로 1픽셀씩 선을 두껍게 만듭니다.
        # (만약 더 굵게 만들고 싶다면 3을 5로 바꾸시면 됩니다)
        img = Image.merge('RGBA', (r, g, b, a))
    # -----------------------------------

    # 엑셀에 들어갈 기본 크기(40x40)로 우선 조절
    img = img.resize((40, 40), Image.Resampling.LANCZOS)
    
    # 투명한 여백을 포함한 새 이미지 캔버스 생성 (가로: 기존너비 + 패딩, 세로: 기존높이)
    new_width = 40 + left_pad
    new_img = Image.new('RGBA', (new_width, 40), (255, 255, 255, 0))
    
    # 새 캔버스 위의 (left_pad, 0) 좌표에 원래 이미지를 붙여넣기
    new_img.paste(img, (left_pad, 0))
    
    out_path = os.path.join(signs_dir, out_filename)
    new_img.save(out_path, 'PNG')
    
    thick_msg = " (굵기 강화 적용)" if make_thick else ""
    print(f'저장 성공: {out_filename} (왼쪽 여백 {left_pad}px 추가){thick_msg}')


if __name__ == "__main__":
    print("도장 이미지 처리 작업을 시작합니다...")
    
    # 유상훈, 주진철: 연필이나 볼펜으로 쓴 얇은 서명이라 너무 희미하므로 make_thick=True 를 주어 아주 까맣고 굵게 만듭니다!
    process_signature('유상훈.png', 15, '유상훈_padded.png', make_thick=True)
    process_signature('주진철.png', 15, '주진철_padded.png', make_thick=True)
    
    # 강신태: 원래 빨간색 도장 이미지이므로 굵기 조절 없이 위치(패딩 45px)만 이동시킵니다.
    process_signature('강신태.png', 45, '강신태_padded.png', make_thick=False)
    
    print("모든 작업이 완료되었습니다!")
