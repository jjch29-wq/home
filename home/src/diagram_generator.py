import os
import math
from PIL import Image, ImageDraw, ImageFont

def draw_dashed_line(draw, start, end, fill="orange", width=2, dash_length=10):
    x1, y1 = start
    x2, y2 = end
    
    if x1 == x2:  # 수직선
        for y in range(min(y1, y2), max(y1, y2), dash_length * 2):
            draw.line([(x1, y), (x1, min(y + dash_length, max(y1, y2)))], fill=fill, width=width)
    elif y1 == y2:  # 수평선
        for x in range(min(x1, x2), max(x1, x2), dash_length * 2):
            draw.line([(x, y1), (min(x + dash_length, max(x1, x2)), y1)], fill=fill, width=width)

def draw_arrow(draw, start, end, fill="blue", width=2, arrow_size=8, double=True):
    draw.line([start, end], fill=fill, width=width)
    x1, y1 = start
    x2, y2 = end
    length = math.hypot(x2 - x1, y2 - y1)
    if length == 0: return
    # 끝부분 화살표 머리
    angle = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - arrow_size * math.cos(angle - math.pi/6), y2 - arrow_size * math.sin(angle - math.pi/6))
    p2 = (x2 - arrow_size * math.cos(angle + math.pi/6), y2 - arrow_size * math.sin(angle + math.pi/6))
    draw.polygon([end, p1, p2], fill=fill)
    
    if double:
        # 시작부분 화살표 머리
        angle = math.atan2(y1 - y2, x1 - x2)
        p1 = (x1 - arrow_size * math.cos(angle - math.pi/6), y1 - arrow_size * math.sin(angle - math.pi/6))
        p2 = (x1 - arrow_size * math.cos(angle + math.pi/6), y1 - arrow_size * math.sin(angle + math.pi/6))
        draw.polygon([start, p1, p2], fill=fill)

def get_korean_font(size=14):
    # 윈도우 맑은 고딕 폰트 시도
    font_paths = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/gulim.ttc",
        "C:/Windows/Fonts/batang.ttc"
    ]
    for path in font_paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()

def create_radiation_diagram(top_dist, bottom_dist, left_dist, right_dist, pipe_len, pipe_width, output_path="temp_diagram.png"):
    # 이미지 크기 설정 (범례와 도면 사이 간격 확보를 위해 세로 크기 증가)
    img_w, img_h = 1000, 650
    img = Image.new('RGB', (img_w, img_h), color='white')
    draw = ImageDraw.Draw(img)
    font = get_korean_font(18)
    font_small = get_korean_font(14)
    font_large = get_korean_font(24)

    # 좌표 정의
    center_x = img_w // 2
    center_y = img_h // 2 - 20 # 약간 위로 치우치게
    
    # 1. 방사선 관리구역 (외곽 노란 점선)
    # 파이프 크기를 대략 500x100으로 가정
    p_w, p_h = 500, 100
    p_left, p_top = center_x - p_w//2, center_y - p_h//2
    p_right, p_bottom = center_x + p_w//2, center_y + p_h//2
    
    # 관리구역 크기
    area_left, area_top = p_left - 150, p_top - 150
    area_right, area_bottom = p_right + 150, p_bottom + 150
    
    # 외곽 점선 그리기
    draw_dashed_line(draw, (area_left, area_top), (area_right, area_top), fill="#f59e0b", width=4, dash_length=15)
    draw_dashed_line(draw, (area_left, area_bottom), (area_right, area_bottom), fill="#f59e0b", width=4, dash_length=15)
    draw_dashed_line(draw, (area_left, area_top), (area_left, area_bottom), fill="#f59e0b", width=4, dash_length=15)
    draw_dashed_line(draw, (area_right, area_top), (area_right, area_bottom), fill="#f59e0b", width=4, dash_length=15)
    
    # 2. 파이프 외곽 박스 및 내부 배관 그리기
    # 외곽 박스 (흰색 바탕, 파란 테두리)
    draw.rectangle([p_left, p_top, p_right, p_bottom], fill="white", outline="#3b82f6", width=2)
    
    # 내부 얇은 배관 바 (노란색/파란색 띠)
    bar_h = 30 # 얇은 배관 높이
    bar_top = center_y - bar_h//2
    bar_bottom = center_y + bar_h//2
    shield_w = 150
    
    # 좌측 배관
    draw.rectangle([p_left, bar_top, center_x - shield_w//2, bar_bottom], fill="#ffedd5", outline="#3b82f6", width=1)
    # 우측 배관
    draw.rectangle([center_x + shield_w//2, bar_top, p_right, bar_bottom], fill="#ffedd5", outline="#3b82f6", width=1)
    # 중앙 납차폐체 (파란 바탕)
    draw.rectangle([center_x - shield_w//2, bar_top, center_x + shield_w//2, bar_bottom], fill="#60a5fa", outline="#3b82f6", width=1)
    
    # 텍스트
    draw.text((p_left + 5, p_bottom - 20), "PIPE", fill="black", font=font_small)
    draw.text((center_x - 35, center_y - 10), "납차폐체", fill="black", font=font)
    
    # 3. 화살표 및 치수 텍스트
    # 상단 거리
    draw_arrow(draw, (center_x, area_top), (center_x, p_top), fill="#60a5fa", width=2)
    draw.text((center_x + 10, area_top + 70), str(top_dist), fill="black", font=font)
    
    # 하단 거리
    draw_arrow(draw, (center_x, p_bottom), (center_x, area_bottom), fill="#60a5fa", width=2)
    draw.text((center_x + 10, p_bottom + 70), str(bottom_dist), fill="black", font=font)
    
    # 좌측 거리
    draw_arrow(draw, (area_left, center_y), (p_left, center_y), fill="#60a5fa", width=2)
    draw.text((area_left + 50, center_y + 10), str(left_dist), fill="black", font=font)
    
    # 우측 거리
    draw_arrow(draw, (p_right, center_y), (area_right, center_y), fill="#60a5fa", width=2)
    draw.text((p_right + 50, center_y + 10), str(right_dist), fill="black", font=font)
    
    # 배관 길이
    draw_arrow(draw, (p_left, p_bottom + 20), (p_right, p_bottom + 20), fill="#60a5fa", width=2)
    draw.text((center_x - 20, p_bottom + 25), str(pipe_len), fill="black", font=font)
    
    # 배관 폭 (우측 안쪽)
    draw_arrow(draw, (p_right - 20, p_top), (p_right - 20, p_bottom), fill="#60a5fa", width=1)
    draw.text((p_right - 70, p_top + 10), str(pipe_width), fill="black", font=font)
    
    # 4. 아이콘 그리기 (경고등, 방사선표지)
    def draw_warning_light(x, y):
        # 빨간 원기둥 느낌
        draw.rectangle([x-10, y-5, x+10, y+10], fill="#1f2937") # 검은 받침대
        draw.ellipse([x-12, y-20, x+12, y], fill="#ef4444") # 빨간 등
        
    def draw_trefoil(cx, cy, size):
        r3 = size
        r2 = size * 0.3
        r1 = size * 0.15
        color = "#e81123" # 붉은색 방사선 로고
        # 3개의 날개 그리기
        for start, end in [(60, 120), (180, 240), (300, 360)]:
            draw.pieslice([cx-r3, cy-r3, cx+r3, cy+r3], start, end, fill=color)
        # 안쪽 갭(노란색으로 덮기)
        draw.ellipse([cx-r2, cy-r2, cx+r2, cy+r2], fill="#fde047")
        # 중앙 원
        draw.ellipse([cx-r1, cy-r1, cx+r1, cy+r1], fill=color)

    def draw_radiation_sign(x, y, label=""):
        # 노란 네모 방사선 표지 배경 (아래쪽을 늘려서 글자가 쏙 들어가게)
        draw.rectangle([x-24, y-25, x+24, y+35], fill="#fde047", outline="#ca8a04", width=1)
        # 진짜 방사선 로고 그리기 (중앙보다 약간 위에)
        draw_trefoil(x, y-6, 12)
        # 아래에 작은 글씨
        draw.text((x-18, y+13), "방사선", fill="black", font=font_small)
        
        # 라벨 (A, B, C, D)
        if label:
            lx, ly = 0, 0
            # 원본 도면과 동일하게 A, B는 우측, C는 좌측, D는 우측으로 바짝 붙여 배치
            if label == "A": lx, ly = x+25, y-20
            elif label == "B": lx, ly = x+25, y-10
            elif label == "C": lx, ly = x-55, y-15
            elif label == "D": lx, ly = x+25, y-15
            draw.rectangle([lx, ly, lx+30, ly+30], fill="#fef08a", outline="black", width=1)
            draw.text((lx+8, ly+3), label, fill="black", font=font_large)

    # 경고등 4모서리
    draw_warning_light(area_left, area_top)
    draw_warning_light(area_right, area_top)
    draw_warning_light(area_left, area_bottom)
    draw_warning_light(area_right, area_bottom)
    
    # 방사선 표지 4면 중앙
    draw_radiation_sign(center_x, area_top, "A")
    draw_radiation_sign(center_x, area_bottom, "B")
    draw_radiation_sign(area_left, center_y, "C")
    draw_radiation_sign(area_right, center_y, "D")
    
    # 5. 범례 (하단)
    leg_y = img_h - 60
    legend_y = leg_y
    draw_warning_light(center_x - 200, leg_y)
    draw.text((center_x - 170, leg_y - 15), "경고등", fill="black", font=font)
    
    draw_radiation_sign(center_x - 50, leg_y)
    draw.text((center_x - 10, leg_y - 15), "방사선표지", fill="black", font=font)
    
    draw_dashed_line(draw, (center_x + 110, leg_y), (center_x + 180, leg_y), fill="#f59e0b", width=3, dash_length=10)
    draw.text((center_x + 190, leg_y - 15), "방사선관리구역", fill="black", font=font)
    
    img.save(output_path, quality=95)
    return output_path

def create_side_diagram(left_dist, right_dist, trench_width, output_path="temp_side_diagram.png"):
    img_w, img_h = 1000, 700
    img = Image.new('RGB', (img_w, img_h), color='white')
    draw = ImageDraw.Draw(img)
    font = get_korean_font(16)
    font_small = get_korean_font(12)
    font_large = get_korean_font(20)
    font_math = get_korean_font(18)
    font_title = get_korean_font(22)
    
    # 타이틀
    draw.text((30, 20), "다. 방사선관리구역 및 감시구역 횡방향 측면도(매설 배관)", fill="black", font=font_title)

    # 1. 흙 (초록색 배경)
    ground_y = 200
    draw.rectangle([20, ground_y, img_w - 20, 420], fill="#7cb342", outline="black")
    draw.text((img_w - 50, 390), "흙", fill="black", font=font)

    # 2. 트렌치 파기 (사다리꼴)
    trench_top_w = 280
    trench_bot_w = 260
    trench_h = 180
    trench_bot_y = ground_y + trench_h # 380
    
    cx = img_w // 2 - 50
    
    trench_poly = [
        (cx - trench_top_w//2, ground_y),
        (cx + trench_top_w//2, ground_y),
        (cx + trench_bot_w//2, trench_bot_y),
        (cx - trench_bot_w//2, trench_bot_y)
    ]
    draw.polygon(trench_poly, fill="white", outline="#3b82f6", width=2)
    draw.line([(cx - trench_top_w//2, ground_y), (cx + trench_top_w//2, ground_y)], fill="white", width=4)
    draw.line([(cx - trench_top_w//2, ground_y), (cx + trench_top_w//2, ground_y)], fill="#3b82f6", width=1) 
    
    # 3. 배관과 차폐체
    pipe_radius = 25
    pipe_cy = trench_bot_y - 15 - pipe_radius # 340
    pipe_y_top = pipe_cy - pipe_radius
    pipe_y_bot = pipe_cy + pipe_radius
    
    draw.ellipse([cx - pipe_radius, pipe_y_top, cx + pipe_radius, pipe_y_bot], fill="white", outline="#3b82f6", width=2)
    draw.text((cx - 12, pipe_cy - 8), "150", fill="black", font=font_small)
    draw_arrow(draw, (cx, pipe_y_top), (cx, pipe_y_bot), fill="#f59e0b", width=1, arrow_size=4, double=True)
    
    # 파란 반원
    draw.arc([cx - pipe_radius - 8, pipe_y_top - 8, cx + pipe_radius + 8, pipe_y_bot + 8], 180, 360, fill="#3b82f6", width=4)
    
    # 빨간 점 (콜리메이터)
    collimator_y = pipe_y_top # 315
    draw.ellipse([cx - 4, collimator_y - 4, cx + 4, collimator_y + 4], fill="red")
    
    # 4. 좌표 및 수학적 값
    try:
        W = float(trench_width) / 2 # 500
    except:
        W = 500
        
    try:
        person_dist = float(right_dist) # 2000
    except:
        person_dist = 2000
        
    person_x_math = W + person_dist # 2500
    pipe_depth = 1500
    person_height = 2000
    
    slope = (pipe_depth + person_height) / person_x_math if person_x_math > 0 else 1.4
    angle_rad = math.atan(slope)
    angle_deg = math.degrees(angle_rad)
    
    A_math = slope * W # 700
    B_math = pipe_depth - A_math # 800
    C_math = B_math / math.sin(angle_rad) if angle_rad > 0 else 0
    
    # 화면 픽셀 좌표 매핑 (스케매틱 뷰)
    wall_x = cx + trench_top_w//2 - 5
    person_x = wall_x + 350
    person_y = ground_y - 100
    wall_y = collimator_y - 40
    
    # 파란 선 (각도 선)
    draw_arrow(draw, (cx, collimator_y), (wall_x + 50, wall_y - 20), fill="#3b82f6", width=1, double=False)
    
    # 빨간 선 (광선)
    draw_arrow(draw, (cx, collimator_y), (wall_x, wall_y), fill="red", width=2, double=False)
    draw_arrow(draw, (wall_x, wall_y), (person_x, person_y + 10), fill="red", width=2, double=False)
    
    # 수평 파란 선
    draw_arrow(draw, (cx, collimator_y), (wall_x, collimator_y), fill="#3b82f6", width=1, double=False)
    
    # 각도 표시
    draw.text((cx + 25, collimator_y - 20), f"{angle_deg:.2f}", fill="black", font=font_small)
    
    # A, B, C 빨간 텍스트
    # A선
    draw_arrow(draw, (wall_x, collimator_y), (wall_x, wall_y), fill="red", width=1, double=True)
    draw.text((wall_x + 5, (collimator_y + wall_y)//2 - 10), "A", fill="red", font=font_small)
    # B선
    draw_arrow(draw, (wall_x, wall_y), (wall_x, ground_y), fill="red", width=1, double=True)
    draw.text((wall_x - 15, (wall_y + ground_y)//2 - 10), "B", fill="red", font=font_small)
    # C텍스트
    draw.text(((wall_x + person_x)//2 - 10, (wall_y + person_y)//2 + 10), "C", fill="red", font=font_small)
    
    # 5. 사람 아이콘
    px = person_x - 10
    py = person_y
    draw.ellipse([px+5, py-20, px+15, py-10], fill="black") # 머리
    draw.rectangle([px, py-10, px+20, py+20], fill="black") # 몸통
    draw.rectangle([px+5, py+20, px+10, py+50], fill="black") # 다리
    draw.rectangle([px+10, py+20, px+15, py+50], fill="black") # 다리
    
    # 6. 치수선
    # 500
    draw_arrow(draw, (cx, collimator_y + 15), (wall_x, collimator_y + 15), fill="#3b82f6", width=1, double=True)
    draw.text((cx + 40, collimator_y + 20), "500", fill="black", font=font_small)
    
    # 2500
    draw_arrow(draw, (cx, trench_bot_y + 15), (person_x, trench_bot_y + 15), fill="#3b82f6", width=1, double=True)
    draw.text(((cx + person_x)//2, trench_bot_y + 20), "2500", fill="black", font=font_small)
    
    # 2000 (우측 흙 안쪽 위)
    draw_arrow(draw, (wall_x, ground_y - 15), (person_x, ground_y - 15), fill="#3b82f6", width=1, double=True)
    draw.text(((wall_x + person_x)//2, ground_y - 30), str(right_dist), fill="black", font=font_small)
    
    # 1000
    draw_arrow(draw, (cx - trench_bot_w//2, trench_bot_y + 10), (cx + trench_bot_w//2, trench_bot_y + 10), fill="#3b82f6", width=1, double=True)
    draw.text((cx - 15, trench_bot_y + 15), "1000", fill="black", font=font_small)
    
    # 100 (배관 아래)
    draw_arrow(draw, (cx, pipe_y_bot), (cx, trench_bot_y), fill="#3b82f6", width=1, double=True)
    draw.text((cx + 5, pipe_y_bot), "100", fill="black", font=font_small)
    
    # 좌측 2000
    left_x = 60
    draw_arrow(draw, (left_x, ground_y - 15), (cx, ground_y - 15), fill="#3b82f6", width=1, double=True)
    draw.text(((left_x + cx)//2, ground_y - 30), str(left_dist), fill="black", font=font_small)
    
    # 1500 수직
    draw_arrow(draw, (person_x + 30, ground_y), (person_x + 30, collimator_y), fill="#3b82f6", width=1, double=True)
    draw.text((person_x + 10, (ground_y + collimator_y)//2 - 10), "1500", fill="black", font=font_small)
    
    # 1750 수직
    draw_arrow(draw, (person_x + 70, ground_y), (person_x + 70, trench_bot_y), fill="#3b82f6", width=1, double=True)
    draw.text((person_x + 50, (ground_y + trench_bot_y)//2 - 10), "1750", fill="black", font=font_small)
    
    # 2000 사람 키
    draw_arrow(draw, (person_x + 30, person_y - 20), (person_x + 30, ground_y), fill="#3b82f6", width=1, double=True)
    draw.text((person_x + 35, (person_y + ground_y)//2 - 20), "2000", fill="black", font=font_small)
    
    # 방사선 감시구역 (점선 및 텍스트 박스)
    box_w = 200
    box_h = 30
    box_x = cx - 50
    box_y = 60
    draw.rectangle([box_x, box_y, box_x + box_w, box_y + box_h], outline="red", width=2)
    draw.text((box_x + 40, box_y + 3), "방사선감시구역", fill="red", font=font)
    
    draw_dashed_line(draw, (left_x, box_y + box_h//2), (box_x, box_y + box_h//2), fill="red", width=2, dash_length=8)
    draw_dashed_line(draw, (box_x + box_w, box_y + box_h//2), (person_x + 20, box_y + box_h//2), fill="red", width=2, dash_length=8)
    draw_dashed_line(draw, (left_x, box_y + box_h//2), (left_x, ground_y), fill="red", width=2, dash_length=8)
    
    # 텍스트 단위:mm
    draw.text((left_x + 10, box_y + box_h + 10), "단위:mm", fill="black", font=font_small)
    
    # 콜리메이터 범례
    draw.ellipse([img_w - 150, 440, img_w - 140, 450], fill="red")
    draw.text((img_w - 130, 435), "콜리메이터", fill="black", font=font)
    
    # 7. 표 그리기
    table_y = 480
    table_margin = 100
    col1_w = 300
    col2_w = img_w - 2 * table_margin - col1_w
    
    draw.rectangle([table_margin, table_y, img_w - table_margin, table_y + 180], outline="black", width=2)
    draw.line([(table_margin, table_y + 40), (img_w - table_margin, table_y + 40)], fill="black", width=2)
    draw.line([(table_margin + col1_w, table_y), (table_margin + col1_w, table_y + 180)], fill="black", width=2)
    
    draw.text((table_margin + col1_w//2 - 20, table_y + 10), "구분", fill="black", font=font_large)
    draw.text((table_margin + col1_w + col2_w//2 - 30, table_y + 10), "계산식", fill="black", font=font_large)
    
    draw.text((table_margin + col1_w//2 - 90, table_y + 90), "토양 차폐두께 평가", fill="black", font=font_math)
    
    eq_x = table_margin + col1_w + 50
    draw.text((eq_x, table_y + 60), f"A : tan {angle_deg:.2f}° × {W:.0f} = {A_math:.0f}mm", fill="black", font=font_math)
    draw.text((eq_x, table_y + 100), f"B : {pipe_depth:,} - {A_math:.0f} = {B_math:.0f}mm", fill="black", font=font_math)
    
    draw.text((eq_x, table_y + 140), f"C :", fill="black", font=font_math)
    c_eq_x = eq_x + 40
    draw.text((c_eq_x + 30, table_y + 130), f"{B_math:.0f}", fill="black", font=font_math)
    draw.text((c_eq_x, table_y + 155), f"sin {angle_deg:.2f}°", fill="black", font=font_math)
    draw.line([(c_eq_x, table_y + 152), (c_eq_x + 100, table_y + 152)], fill="black", width=2)
    draw.text((c_eq_x + 110, table_y + 140), f"= {C_math:.0f}mm", fill="black", font=font_math)
    
    img.save(output_path, quality=95)
    return output_path

if __name__ == "__main__":
    # 테스트 실행
    create_radiation_diagram("2000", "2000", "2000", "2000", "10000", "1000", "test_diagram.jpg")
    create_side_diagram("2000", "2000", "1000", "test_side_diagram.jpg")
    print("test_diagram.jpg, test_side_diagram.jpg 생성 완료")
