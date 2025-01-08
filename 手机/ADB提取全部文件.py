import subprocess
import os

# 定义保存APP的路径
save_path = "F:/system/Pictures/新建文件夹"
if not os.path.exists(save_path):
    os.makedirs(save_path)


# 获取已安装的APP列表
def get_installed_apps():
    result = subprocess.run(['adb', 'shell', 'pm', 'list', 'packages'], stdout=subprocess.PIPE)
    packages = result.stdout.decode('utf-8').splitlines()
    packages = [pkg.split(':')[1] for pkg in packages]
    return packages


# 提取每个APP并保存到指定目录
def pull_apps(packages):
    for package in packages:
        apk_path_command = f"adb shell pm path {package}"
        result = subprocess.run(apk_path_command.split(), stdout=subprocess.PIPE)
        apk_path = result.stdout.decode('utf-8').strip().split(":")[1]

        local_path = os.path.join(save_path, f"{package}.apk")
        pull_command = f"adb pull {apk_path} {local_path}"
        subprocess.run(pull_command.split())

        print(f"Pulled {package} to {local_path}")


# 主程序
if __name__ == "__main__":
    packages = get_installed_apps()
    pull_apps(packages)
    print("所有APP已成功提取。")
