import os
from PIL import Image, ImageDraw, ImageFont
import math

def get_korean_font(size=14):
    font_paths = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/gulim.ttc",
        "C:/Windows/Fonts/batang.ttc"
    ]
    for path in font_paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()

def draw_arrow(draw, start, end, fill="blue", width=2, arrow_size=8):
    draw.line([start, end], fill=fill, width=width)
    x1, y1 = start
    x2, y2 = end
    length = math.dist(start, end)
    if length == 0: return
    angle = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - arrow_size * math.cos(angle - math.pi/6), y2 - arrow_size * math.sin(angle - math.pi/6))
    p2 = (x2 - arrow_size * math.cos(angle + math.pi/6), y2 - arrow_size * math.sin(angle + math.pi/6))
    draw.polygon([end, p1, p2], fill=fill)
    
def draw_double_arrow(draw, start, end, fill="blue", width=1, arrow_size=6):
    draw.line([start, end], fill=fill, width=width)
    x1, y1 = start
    x2, y2 = end
    length = math.dist(start, end)
    if length == 0: return
    # end arrow
    angle = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - arrow_size * math.cos(angle - math.pi/6), y2 - arrow_size * math.sin(angle - math.pi/6))
    p2 = (x2 - arrow_size * math.cos(angle + math.pi/6), y2 - arrow_size * math.sin(angle + math.pi/6))
    draw.polygon([end, p1, p2], fill=fill)
    # start arrow
    angle = math.atan2(y1 - y2, x1 - x2)
    p1 = (x1 - arrow_size * math.cos(angle - math.pi/6), y1 - arrow_size * math.sin(angle - math.pi/6))
    p2 = (x1 - arrow_size * math.cos(angle + math.pi/6), y1 - arrow_size * math.sin(angle + math.pi/6))
    draw.polygon([start, p1, p2], fill=fill)

def draw_dashed_line(draw, start, end, fill="orange", width=2, dash_length=10):
    x1, y1 = start
    x2, y2 = end
    length = math.dist(start, end)
    if length == 0: return
    dx = (x2 - x1) / length
    dy = (y2 - y1) / length
    for i in range(0, int(length), dash_length * 2):
        d_start = (x1 + dx * i, y1 + dy * i)
        end_len = min(i + dash_length, length)
        d_end = (x1 + dx * end_len, y1 + dy * end_len)
        draw.line([d_start, d_end], fill=fill, width=width)

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
    draw_double_arrow(draw, (cx, pipe_y_top), (cx, pipe_y_bot), fill="#f59e0b", width=1, arrow_size=4)
    
    # 파란 반원
    draw.arc([cx - pipe_radius - 8, pipe_y_top - 8, cx + pipe_radius + 8, pipe_y_bot + 8], 180, 360, fill="#3b82f6", width=4)
    
    # 빨간 점 (콜리메이터)
    collimator_y = pipe_y_top # 315
    draw.ellipse([cx - 4, collimator_y - 4, cx + 4, collimator_y + 4], fill="red")
    
    # 4. 좌표 및 수학적 값
    W = 500
    person_dist = int(right_dist) # 2000
    person_x_math = W + person_dist # 2500
    pipe_depth = 1500
    person_height = 2000
    
    slope = (pipe_depth + person_height) / person_x_math # 3500 / 2500 = 1.4
    angle_rad = math.atan(slope)
    angle_deg = math.degrees(angle_rad)
    
    A_math = slope * W # 700
    B_math = pipe_depth - A_math # 800
    C_math = B_math / math.sin(angle_rad) # 983
    
    # 화면 픽셀 좌표 매핑 (스케매틱 뷰)
    # 우측 벽 X
    wall_x = cx + trench_top_w//2 - 5 # 약간 안쪽
    # 사람 X
    person_x = wall_x + 350
    # 사람 Y
    person_y = ground_y - 100
    
    # 광선: (cx, collimator_y) -> (wall_x, wall_y) -> (person_x, person_y)
    # wall_y 를 비율에 맞게
    # 전체 높이: collimator_y -> person_y. (315 -> 100 = 215 px)
    # A_math (700) : B_math+person_height (800+2000=2800)
    # 벽 교차점은 위에서부터 B+person_height. 
    # 대략적으로 눈대중으로 그립니다.
    wall_y = collimator_y - 40 # 275
    
    # 파란 선 (각도 선)
    draw_arrow(draw, (cx, collimator_y), (wall_x + 50, wall_y - 20), fill="#3b82f6", width=1)
    
    # 빨간 선 (광선)
    draw_arrow(draw, (cx, collimator_y), (wall_x, wall_y), fill="red", width=2)
    draw_arrow(draw, (wall_x, wall_y), (person_x, person_y + 10), fill="red", width=2)
    
    # 수평 파란 선
    draw_arrow(draw, (cx, collimator_y), (wall_x, collimator_y), fill="#3b82f6", width=1)
    
    # 각도 표시
    draw.text((cx + 25, collimator_y - 20), f"{angle_deg:.2f}", fill="black", font=font_small)
    
    # A, B, C 빨간 텍스트
    # A선
    draw_double_arrow(draw, (wall_x, collimator_y), (wall_x, wall_y), fill="red", width=1)
    draw.text((wall_x + 5, (collimator_y + wall_y)//2 - 10), "A", fill="red", font=font_small)
    # B선
    draw_double_arrow(draw, (wall_x, wall_y), (wall_x, ground_y), fill="red", width=1)
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
    draw_double_arrow(draw, (cx, collimator_y + 15), (wall_x, collimator_y + 15), fill="#3b82f6", width=1)
    draw.text((cx + 40, collimator_y + 20), "500", fill="black", font=font_small)
    
    # 2500
    draw_double_arrow(draw, (cx, trench_bot_y + 15), (person_x, trench_bot_y + 15), fill="#3b82f6", width=1)
    draw.text(((cx + person_x)//2, trench_bot_y + 20), "2500", fill="black", font=font_small)
    
    # 2000 (우측 흙 안쪽 위)
    draw_double_arrow(draw, (wall_x, ground_y - 15), (person_x, ground_y - 15), fill="#3b82f6", width=1)
    draw.text(((wall_x + person_x)//2, ground_y - 30), str(right_dist), fill="black", font=font_small)
    
    # 1000
    draw_double_arrow(draw, (cx - trench_bot_w//2, trench_bot_y + 10), (cx + trench_bot_w//2, trench_bot_y + 10), fill="#3b82f6", width=1)
    draw.text((cx - 15, trench_bot_y + 15), "1000", fill="black", font=font_small)
    
    # 100 (배관 아래)
    draw_double_arrow(draw, (cx, pipe_y_bot), (cx, trench_bot_y), fill="#3b82f6", width=1)
    draw.text((cx + 5, pipe_y_bot), "100", fill="black", font=font_small)
    
    # 좌측 2000
    left_x = 60
    draw_double_arrow(draw, (left_x, ground_y - 15), (cx, ground_y - 15), fill="#3b82f6", width=1)
    draw.text(((left_x + cx)//2, ground_y - 30), str(left_dist), fill="black", font=font_small)
    
    # 1500 수직
    draw_double_arrow(draw, (person_x + 30, ground_y), (person_x + 30, collimator_y), fill="#3b82f6", width=1)
    draw.text((person_x + 10, (ground_y + collimator_y)//2 - 10), "1500", fill="black", font=font_small)
    
    # 1750 수직
    draw_double_arrow(draw, (person_x + 70, ground_y), (person_x + 70, trench_bot_y), fill="#3b82f6", width=1)
    draw.text((person_x + 50, (ground_y + trench_bot_y)//2 - 10), "1750", fill="black", font=font_small)
    
    # 2000 사람 키
    draw_double_arrow(draw, (person_x + 30, person_y - 20), (person_x + 30, ground_y), fill="#3b82f6", width=1)
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
    create_side_diagram(2000, 2000, 1000, "test_side.jpg")
    print("test_side.jpg created")
