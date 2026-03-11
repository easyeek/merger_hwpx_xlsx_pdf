from PIL import Image
import os

source_img = r'C:\Users\user\.gemini\antigravity\brain\289e0dff-683c-4ea4-9d24-63954f931f6c\easyeek_pixel_icon_1773242511537.png'
target_ico = r'e:\00_Project\05_merge(hwpx, xlsx)\app_icon.ico'

if os.path.exists(source_img):
    img = Image.open(source_img)
    # Windows 아이콘 표준 사이즈들
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(target_ico, format='ICO', sizes=icon_sizes)
    print(f"Successfully converted {source_img} to {target_ico}")
else:
    print(f"Error: Source image not found at {source_img}")
