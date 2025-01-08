import os

# 获取 PyCharm 配置文件路径
pycharm_config_path = os.path.expanduser("~/.PyCharm<version>/config/pycharm.vmoptions")

# 检查文件是否存在
if os.path.exists(pycharm_config_path):
    with open(pycharm_config_path, 'r') as file:
        lines = file.readlines()

    # 修改内存设置
    with open(pycharm_config_path, 'w') as file:
        for line in lines:
            if line.startswith("-Xms"):
                file.write("-Xms512m\n")
            elif line.startswith("-Xmx"):
                file.write("-Xmx4096m\n")
            else:
                file.write(line)
    print("PyCharm 内存设置已更改为 4GB。")
else:
    print(f"未找到配置文件：{pycharm_config_path}")
