from PIL import Image

try:
    img = Image.open('static/logo.png')
    w, h = img.size
    print(f"Size: {w}x{h}")
    
    # We assume the "f" icon is a square on the left side of the logo.
    # So we crop a box (0, 0, h, h) which is a square matching the height.
    icon = img.crop((0, 0, h, h))
    icon.save('static/favicon.png')
    print("Cropped and saved to favicon.png")
except Exception as e:
    print("Error:", e)
