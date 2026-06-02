import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
import os

def thicken_signature(image_path):
    # 이미지 불러오기 (알파 채널 포함)
    # 한글 경로 지원을 위해 numpy로 읽어서 디코딩
    stream = open(image_path, "rb")
    bytes_arr = bytearray(stream.read())
    numpyarray = np.asarray(bytes_arr, dtype=np.uint8)
    img = cv2.imdecode(numpyarray, cv2.IMREAD_UNCHANGED)
    
    if img is None:
        print("이미지를 불러올 수 없습니다.")
        return

    # 그레이스케일 변환을 통해 글씨 부분 찾기
    if len(img.shape) == 3 and img.shape[2] == 4:
        # 투명 배경(알파 채널)이 있는 경우
        bgr = img[:, :, :3]
        alpha = img[:, :, 3]
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        # 투명한 부분은 흰색 배경으로 취급
        gray[alpha == 0] = 255
    else:
        # 배경이 이미 있는 경우 (RGB 또는 Gray)
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            gray = img

    # 이진화 (검은 글씨를 흰색(255)으로, 흰 배경을 검은색(0)으로 반전)
    # 적응형 스레시홀드 또는 고정 임계값 적용
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    # 굵기 조절 옵션들 (Kernel 크기가 클수록 더 굵어짐)
    kernels = [
        ("1_약간굵게", np.ones((2, 2), np.uint8), 1),
        ("2_중간굵게", np.ones((3, 3), np.uint8), 1),
        ("3_매우굵게", np.ones((4, 4), np.uint8), 1),
        ("4_가장굵게", np.ones((5, 5), np.uint8), 1)
    ]

    base_dir, ext = os.path.splitext(image_path)
    
    # 원본 파일명 추출
    filename = os.path.basename(base_dir)
    dir_path = os.path.dirname(base_dir)

    for name, kernel, iters in kernels:
        # 팽창(Dilation) 적용하여 글씨 영역 확장
        dilated = cv2.dilate(binary, kernel, iterations=iters)

        # 결과 이미지 생성 (투명 배경 + 검은색 글씨)
        # B=0, G=0, R=0 (검은색)
        out_img = np.zeros((img.shape[0], img.shape[1], 4), dtype=np.uint8)
        
        # 알파 채널에 팽창된 마스크 적용 (글씨 있는 곳만 불투명하게 255)
        out_img[:, :, 3] = dilated
        
        # 한글 경로 저장을 위한 인코딩
        out_path = os.path.join(dir_path, f"{filename}_{name}.png")
        result, encoded_img = cv2.imencode('.png', out_img)
        if result:
            with open(out_path, mode='w+b') as f:
                encoded_img.tofile(f)
            print(f"저장 완료: {out_path}")

def main():
    root = tk.Tk()
    root.withdraw()
    # 최상단에 띄우기
    root.attributes('-topmost', True)
    
    print("싸인 이미지 파일을 선택하는 창이 뜹니다...")
    file_path = filedialog.askopenfilename(
        title="굵기를 조절할 싸인 이미지를 선택하세요", 
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")]
    )

    if file_path:
        print(f"선택된 파일: {file_path}")
        thicken_signature(file_path)
        print("\n작업이 완료되었습니다! 원본 이미지와 같은 폴더에 굵기별 투명 배경 파일이 생성되었습니다.")
    else:
        print("파일이 선택되지 않았습니다.")

if __name__ == "__main__":
    main()
