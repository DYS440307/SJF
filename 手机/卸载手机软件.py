import subprocess


def execute_adb_command(command):
    result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout, result.stderr


def read_packages_from_file(file_path):
    with open(file_path, 'r') as file:
        packages = file.readlines()
    return [pkg.strip() for pkg in packages]


def uninstall_package(package):
    command = f'adb shell pm uninstall --user 0 {package}'
    stdout, stderr = execute_adb_command(command)
    return stdout, stderr


def freeze_package(package):
    command = f'adb shell pm disable-user --user 0 {package}'
    stdout, stderr = execute_adb_command(command)
    return stdout, stderr


def main():
    packages_file = input("请输入文件路径: ")
    packages = read_packages_from_file(packages_file)

    for package in packages:
        print(f'Uninstalling package: {package}')
        stdout, stderr = uninstall_package(package)
        if "Failure" in stderr:
            print(f'Failed to uninstall {package}. Attempting to freeze it.')
            stdout, stderr = freeze_package(package)
            if "Error" in stderr:
                print(f'Failed to freeze {package}. Error: {stderr}')
            else:
                print(f'Successfully froze {package}')
        else:
            print(f'Successfully uninstalled {package}')


if __name__ == '__main__':
    main()
