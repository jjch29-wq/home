import os
from PIL import Image, ImageDraw, ImageFont

base_dir = r"C:\Users\-\PMI\home\resources"
logo_path = os.path.join(base_dir, "logo.png")
out_path = os.path.join(base_dir, "logo_joongang_generated.png")

try:
    # 1. Open the original logo (first image in the user's prompt without text)
    img = Image.open(logo_path).convert("RGBA")
    
    # 2. The second image has text added to the right. 
    # Let's create a new image that is wider to accommodate the text.
    width, height = img.size
    
    # Estimate width needed for "중앙지사" (4 characters)
    # We will use a bold font if possible, Malgun Gothic
    font_path = "C:/Windows/Fonts/malgun.ttf" # standard on windows
    try:
        font = ImageFont.truetype(font_path, int(height * 0.45))
    except:
        font = ImageFont.load_default()
        
    text = "중앙지사"
    
    # get text bounding box
    bbox = font.getbbox(text)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    
    # Create new image with extra width
    # Give some padding between logo and text
    padding = 20
    new_width = width + text_w + padding + 20
    new_img = Image.new("RGBA", (new_width, height), (255, 255, 255, 0)) # transparent background
    
    # Paste the original logo
    new_img.paste(img, (0, 0))
    
    # Draw the text
    draw = ImageDraw.Draw(new_img)
    
    # Calculate vertical position to center the text
    # Looking at the user's "동탄지사" image, the text is centered vertically with the logo text
    text_y = (height - text_h) // 2 - int(height * 0.05) # slight adjustment
    text_x = width + padding
    
    draw.text((text_x, text_y), text, font=font, fill=(0, 0, 0, 255))
    
    # Save the result
    new_img.save(out_path)
    print(f"Successfully generated: {out_path}")
    
except Exception as e:
    print(f"Error: {e}")
