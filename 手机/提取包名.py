import subprocess


def get_installed_packages():
    # 运行adb命令获取已安装的包名列表
    result = subprocess.run(['adb', 'shell', 'pm', 'list packages'], stdout=subprocess.PIPE, text=True)

    # 检查adb命令是否成功执行
    if result.returncode != 0:
        print("Failed to execute adb command. Make sure your device is connected and USB debugging is enabled.")
        return []

    # 解析输出，提取包名
    packages = result.stdout.splitlines()
    package_names = [pkg.split(":")[1] for pkg in packages]
    return package_names


def save_packages_to_file(package_names, file_path):
    try:
        with open(file_path, 'w') as file:
            for pkg in package_names:
                file.write(pkg + '\n')
        print(f"Package names saved to {file_path}")
    except Exception as e:
        print(f"Failed to save package names to file: {e}")


def main():
    package_names = get_installed_packages()
    if package_names:
        file_path = input("Enter the file path to save the package names: ")
        save_packages_to_file(package_names, file_path)
    else:
        print("No packages found or failed to retrieve packages.")


if __name__ == '__main__':
    main()
