import pyautogui
import easyocr
import cv2
import numpy as np

# 初始化 EasyOCR 识别器，支持中文和英文
reader = easyocr.Reader(['en', 'ch_sim'])  # 'ch_sim' 是简体中文，'ch_tra' 是繁体中文

# 定义屏幕截图区域
left, top, width, height = 2, 70, 316-2, 1032-70

# 截取屏幕指定区域
screenshot = pyautogui.screenshot(region=(left, top, width, height))
screenshot = np.array(screenshot)
screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

# 使用 EasyOCR 进行文字识别
results = reader.readtext(screenshot)

# 打印识别到的所有文字及其位置
for result in results:
    text = result[1]  # 识别出的文本
    print(f"识别到的文字: {text}")
    # 获取文字框的位置 (左上角和右下角)
    (x1, y1), (x2, y2) = result[0]
    # 计算文字中心点坐标
    x = (x1 + x2) / 2 + left
    y = (y1 + y2) / 2 + top
    print(f"文字位置: ({x}, {y})")

    # 双击该文字的位置
    pyautogui.moveTo(x, y)
    pyautogui.doubleClick()

