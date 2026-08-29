import os
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

def draw_arrow(draw, start, end, fill="blue", width=2):
    draw.line([start, end], fill=fill, width=width)
    # 화살표 머리 그리기 로직은 간단히 생략하거나 꺾인 선으로 추가 가능
    x1, y1 = start
    x2, y2 = end
    
    # 끝부분 화살표 머리
    if x1 == x2: # 수직
        if y1 < y2: # 아래로
            draw.polygon([(x2-5, y2-10), (x2+5, y2-10), (x2, y2)], fill=fill)
            draw.polygon([(x1-5, y1+10), (x1+5, y1+10), (x1, y1)], fill=fill) # 양방향
        else: # 위로
            draw.polygon([(x2-5, y2+10), (x2+5, y2+10), (x2, y2)], fill=fill)
            draw.polygon([(x1-5, y1-10), (x1+5, y1-10), (x1, y1)], fill=fill)
    elif y1 == y2: # 수평
        if x1 < x2: # 오른쪽으로
            draw.polygon([(x2-10, y2-5), (x2-10, y2+5), (x2, y2)], fill=fill)
            draw.polygon([(x1+10, y1-5), (x1+10, y1+5), (x1, y1)], fill=fill)
        else:
            draw.polygon([(x2+10, y2-5), (x2+10, y2+5), (x2, y2)], fill=fill)
            draw.polygon([(x1-10, y1-5), (x1-10, y1+5), (x1, y1)], fill=fill)

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
    img_w, img_h = 1000, 560
    img = Image.new('RGB', (img_w, img_h), color='white')
    draw = ImageDraw.Draw(img)
    font = get_korean_font(18)
    font_small = get_korean_font(14)
    font_large = get_korean_font(24)

    # 1. 흙 (초록색 배경)
    ground_y = 150
    draw.rectangle([20, ground_y, img_w - 20, img_h - 20], fill="#7cb342", outline="black")
    draw.text((img_w - 150, img_h - 150), "흙", fill="black", font=font)

    # 2. 트렌치 파기 (사다리꼴)
    trench_top_w = 400
    trench_bot_w = 350
    trench_h = 300
    trench_bot_y = ground_y + trench_h
    
    cx = img_w // 2
    
    trench_poly = [
        (cx - trench_top_w//2, ground_y),
        (cx + trench_top_w//2, ground_y),
        (cx + trench_bot_w//2, trench_bot_y),
        (cx - trench_bot_w//2, trench_bot_y)
    ]
    draw.polygon(trench_poly, fill="white", outline="#3b82f6", width=2)
    # 덮인 윗 선 지우기 (사다리꼴 그릴때 윗선이 그려짐)
    draw.line([(cx - trench_top_w//2, ground_y), (cx + trench_top_w//2, ground_y)], fill="white", width=4)
    draw.line([(cx - trench_top_w//2 - 2, ground_y), (cx + trench_top_w//2 + 2, ground_y)], fill="#3b82f6", width=1) # 얇게 다시 그어주기
    
    # 3. 배관과 차폐체
    pipe_y_bot = trench_bot_y - 40
    pipe_y_top = pipe_y_bot - 50
    # 배관 좌측 (노란색)
    draw.polygon([(cx - trench_bot_w//2 - 4, pipe_y_top), (cx - 80, pipe_y_top), (cx - 80, pipe_y_bot), (cx - trench_bot_w//2 + 5, pipe_y_bot)], fill="#fef08a", outline="#3b82f6", width=1)
    draw.text((cx - 150, pipe_y_top + 15), "", fill="black", font=font_small)
    # 배관 우측 (노란색)
    draw.polygon([(cx + 80, pipe_y_top), (cx + trench_bot_w//2 + 4, pipe_y_top), (cx + trench_bot_w//2 - 5, pipe_y_bot), (cx + 80, pipe_y_bot)], fill="#fef08a", outline="#3b82f6", width=1)
    draw.text((cx + 100, pipe_y_top + 15), "150", fill="black", font=font_small)
    
    # 중앙 납차폐체 (파란색, 모서리 둥글게)
    draw.rounded_rectangle([cx - 80, pipe_y_top - 5, cx + 80, pipe_y_bot + 5], radius=10, fill="#60a5fa", outline="#3b82f6", width=1)
    draw.text((cx - 35, pipe_y_top + 15), "납차폐체", fill="black", font=font)

    # 4. 방사선 관리구역 박스 및 점선 (상단)
    box_w, box_h = 250, 40
    box_top = 30
    draw.rectangle([cx - box_w//2, box_top, cx + box_w//2, box_top + box_h], fill="white", outline="#eab308", width=3)
    draw.text((cx - 65, box_top + 10), "방사선관리구역", fill="black", font=font)
    
    # 상단 좌우 점선
    draw_dashed_line(draw, (80, box_top + box_h//2), (cx - box_w//2, box_top + box_h//2), fill="#eab308", width=3, dash_length=15)
    draw_dashed_line(draw, (cx + box_w//2, box_top + box_h//2), (img_w - 20, box_top + box_h//2), fill="#eab308", width=3, dash_length=15)
    # 우측 세로 점선
    draw_dashed_line(draw, (img_w - 20, box_top + box_h//2), (img_w - 20, ground_y), fill="#eab308", width=3, dash_length=15)
    # 화살표 끝부분 장식
    draw.polygon([(80, box_top + box_h//2), (95, box_top + box_h//2 - 8), (95, box_top + box_h//2 + 8)], fill="#eab308")
    draw.polygon([(img_w - 20, box_top + box_h//2), (img_w - 35, box_top + box_h//2 - 8), (img_w - 35, box_top + box_h//2 + 8)], fill="#eab308")
    
    # 5. 사람 아이콘 (좌측)
    px, py = 50, ground_y - 100
    draw.ellipse([px+10, py, px+30, py+20], fill="black") # 머리
    draw.rectangle([px+5, py+20, px+35, py+60], fill="black") # 몸통
    draw.rectangle([px+10, py+60, px+20, py+100], fill="black") # 왼쪽다리
    draw.rectangle([px+20, py+60, px+30, py+100], fill="black") # 오른쪽다리
    draw.ellipse([px-5, py+30, px+15, py+50], fill="black") # 팔1
    draw.ellipse([px+25, py+30, px+45, py+50], fill="black") # 팔2
    
    # 6. 치수선
    # 좌측 이격거리
    draw_arrow(draw, (px+35, ground_y - 20), (cx - trench_top_w//2, ground_y - 20), fill="#60a5fa", width=2)
    draw.text((cx - trench_top_w//2 - 120, ground_y - 45), str(left_dist), fill="black", font=font)
    
    # 우측 이격거리
    draw_arrow(draw, (cx + trench_top_w//2, ground_y - 20), (img_w - 20, ground_y - 20), fill="#60a5fa", width=2)
    draw.text((cx + trench_top_w//2 + 100, ground_y - 45), str(right_dist), fill="black", font=font)
    
    # 하단 트렌치 폭
    draw_arrow(draw, (cx - trench_bot_w//2, trench_bot_y + 15), (cx + trench_bot_w//2, trench_bot_y + 15), fill="#60a5fa", width=2)
    draw.text((cx - 20, trench_bot_y + 25), str(trench_width), fill="black", font=font)
    
    # 배관 하단 공간 (100)
    draw_arrow(draw, (cx, pipe_y_bot), (cx, trench_bot_y), fill="#60a5fa", width=2)
    draw.text((cx + 20, pipe_y_bot + 10), "100", fill="black", font=font_small)
    
    # 배관 윗 공간 (1500)
    draw_arrow(draw, (cx - 100, ground_y), (cx - 100, pipe_y_top - 5), fill="#60a5fa", width=2)
    draw.text((cx - 90, ground_y + 70), "1500", fill="black", font=font_small)

    # 상단 우측 단위
    draw.text((img_w - 110, box_top + box_h + 10), "단위:mm", fill="black", font=font_small)

    img.save(output_path, quality=95)
    return output_path

if __name__ == "__main__":
    # 테스트 실행
    create_radiation_diagram("2000", "2000", "2000", "2000", "10000", "1000", "test_diagram.jpg")
    create_side_diagram("2000", "2000", "1000", "test_side_diagram.jpg")
    print("test_diagram.jpg, test_side_diagram.jpg 생성 완료")
