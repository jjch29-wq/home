import cv2
import numpy as np
import fitz  # PyMuPDF
import os

def process_pdf(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    doc = fitz.open(pdf_path)
    print(f"Total pages: {len(doc)}")
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Render page to high resolution image (matrix with zoom)
        zoom = 3.0  # Increase resolution for better detection
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Convert pixmap to numpy array (RGB)
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4: # RGBA to RGB
            img_data = cv2.cvtColor(img_data, cv2.COLOR_RGBA2RGB)
        
        # Convert to Grayscale
        gray = cv2.cvtColor(img_data, cv2.COLOR_RGB2GRAY)
        
        # Apply blur to reduce noise
        gray_blurred = cv2.medianBlur(gray, 5)
        
        # Detect circles
        # param1: gradient threshold, param2: accumulator threshold (lower = more circles)
        circles = cv2.HoughCircles(
            gray_blurred, 
            cv2.HOUGH_GRADIENT, 
            dp=1, 
            minDist=50, 
            param1=100, 
            param2=30, 
            minRadius=10, 
            maxRadius=500
        )
        
        # Create a white background of the same size
        result_img = np.ones_like(img_data) * 255
        
        if circles is not None:
            circles = np.uint16(np.around(circles))
            print(f"Found {len(circles[0])} circles on page {page_num + 1}")
            
            for i in circles[0, :]:
                center = (i[0], i[1])
                radius = i[2]
                
                # Draw a clean black circle outline
                # thickness=2 or more for visibility
                cv2.circle(result_img, center, radius, (0, 0, 0), 2)
        else:
            print(f"No circles found on page {page_num + 1}")
            
        # Save the result
        output_path = os.path.join(output_folder, f"clean_circles_page_{page_num + 1}.png")
        cv2.imwrite(output_path, cv2.cvtColor(result_img, cv2.COLOR_RGB2BGR))
        print(f"Saved: {output_path}")

if __name__ == "__main__":
    pdf_file = r"C:\Users\-\OneDrive\문서\카카오톡 받은 파일\A1-C001-Rev 1.pdf"
    output_dir = "output_circles"
    process_pdf(pdf_file, output_dir)
