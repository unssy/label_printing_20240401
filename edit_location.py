import pyautogui

while True:
    x, y = pyautogui.position() # 获取鼠标位置
    if 0 <= x <= 3839 and 0 <= y <= 1079: # 检查坐标是否在有效范围内
        try:
            pixel_color = pyautogui.pixel(x, y) # 获取指定位置的像素颜色
            print(f"当前鼠标位置：{x}, {y}，RGB：{pixel_color}")
        except Exception as e:
            print(f"获取像素颜色失败：{e}")
    else:
        print("鼠标位置超出屏幕边界")
