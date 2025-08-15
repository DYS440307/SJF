import pyautogui
import datetime
import time
import webbrowser
import os

# ================== 配置区（方便修改） ==================
TARGET_X, TARGET_Y = -900, 940  # 点击坐标
TARGET_URL = "https://cn.cotaiticketing.com/shows/twice2025.html"
TARGET_CLICK_TIME = datetime.time(9, 11, 59, 700000)  # 点击时间（时, 分, 秒, 微秒）
CHROME_PATHS = [
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
]
# ======================================================


def click_position(x, y):
    """瞬间点击指定坐标"""
    try:
        pyautogui.moveTo(x, y, duration=0)
        pyautogui.click()
        print(f"[{datetime.datetime.now()}] 已点击坐标: ({x}, {y})")
    except Exception as e:
        print(f"点击出错: {e}")


def wait_until(target_time):
    """高精度等待直到目标时间"""
    while True:
        now = datetime.datetime.now()
        diff = (target_time - now).total_seconds()
        if diff <= 0:
            break
        if diff > 0.05:  # 粗等待
            time.sleep(0.02)
        else:  # 精细等待
            time.sleep(diff)


def open_chrome_new_window(url):
    """用Chrome新窗口打开网页"""
    try:
        chrome_path = next((p for p in CHROME_PATHS if os.path.exists(p)), None)
        if chrome_path:
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get('chrome').open(url, new=2, autoraise=True)
            print(f"[{datetime.datetime.now()}] 用Chrome新窗口打开: {url}")
        else:
            webbrowser.open(url, new=2, autoraise=True)
            print(f"[{datetime.datetime.now()}] 未找到Chrome，已用默认浏览器打开: {url}")
    except Exception as e:
        print(f"打开网页出错: {e}")


if __name__ == "__main__":
    today = datetime.date.today()
    target_click_dt = datetime.datetime.combine(today, TARGET_CLICK_TIME)

    if target_click_dt < datetime.datetime.now():
        target_click_dt += datetime.timedelta(days=1)

    target_open_dt = target_click_dt - datetime.timedelta(milliseconds=1)

    print(f"等待到 {target_open_dt} 打开网页...")
    wait_until(target_open_dt)
    open_chrome_new_window(TARGET_URL)

    print(f"等待到 {target_click_dt} 点击...")
    wait_until(target_click_dt)
    click_position(TARGET_X, TARGET_Y)
