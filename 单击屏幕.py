import pyautogui
import datetime
import time


def click_position(x, y):
    """瞬间点击屏幕上指定的(x, y)坐标"""
    try:
        # 瞬间移动鼠标到指定位置（duration=0）
        pyautogui.moveTo(x, y, duration=0)

        # 执行鼠标左键单击
        pyautogui.click()
        print(f"已在 {datetime.datetime.now()} 瞬间点击坐标: ({x}, {y})")
    except Exception as e:
        print(f"点击过程中发生错误: {e}")


def wait_until(target_time):
    """等待直到目标时间"""
    while True:
        now = datetime.datetime.now()
        # 如果当前时间已过目标时间，退出循环
        if now >= target_time:
            break
        # 计算剩余时间，最小等待0.001秒（1毫秒）以减少CPU占用
        sleep_time = (target_time - now).total_seconds()
        if sleep_time > 0:
            time.sleep(min(sleep_time, 0.001))


if __name__ == "__main__":
    # 目标坐标
    target_x, target_y = -900, 940

    # 设置目标时间：北京时间08:52:59.8
    # 获取今天的日期，然后设置具体时间
    today = datetime.date.today()
    target_time = datetime.datetime.combine(
        today,
        datetime.time(8, 52, 59, 700000)  # 800000微秒 = 0.8秒
    )

    # 如果目标时间已过今天，则设置为明天的同一时间
    if target_time < datetime.datetime.now():
        target_time += datetime.timedelta(days=1)

    print(f"等待到 {target_time} 执行点击...")
    wait_until(target_time)

    # 执行点击操作
    click_position(target_x, target_y)
