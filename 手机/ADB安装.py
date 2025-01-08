import os
import requests
import zipfile
import subprocess

# 下载 ADB 工具包
def download_adb(url, save_path):
    response = requests.get(url, stream=True)
    with open(save_path, 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)
    print(f"Downloaded ADB to {save_path}")

# 解压 ADB 工具包
def unzip_file(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"Unzipped ADB to {extract_to}")

# 配置环境变量
def set_env_variable(adb_path):
    # 获取当前的环境变量
    current_path = os.environ['PATH']
    if adb_path not in current_path:
        os.environ['PATH'] = f"{adb_path};{current_path}"
        # 持久化设置
        subprocess.run(['setx', 'PATH', os.environ['PATH']])
        print(f"Set ADB path in environment variables: {adb_path}")
    else:
        print(f"ADB path already in environment variables: {adb_path}")

def main():
    adb_url = "https://dl.google.com/android/repository/platform-tools-latest-windows.zip"
    download_path = "platform-tools-latest-windows.zip"
    extract_path = "C:\\adb"

    download_adb(adb_url, download_path)
    unzip_file(download_path, extract_path)
    adb_path = os.path.join(extract_path, 'platform-tools')
    set_env_variable(adb_path)

    # 验证安装
    result = subprocess.run(['adb', 'version'], capture_output=True, text=True)
    print(result.stdout)

if __name__ == "__main__":
    main()
