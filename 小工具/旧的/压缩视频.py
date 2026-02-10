import subprocess
import json
import math
import os


def get_video_info(input_path):
    """
    获取视频的基本信息（时长、音频码率等）
    :param input_path: 输入视频文件路径
    :return: 包含时长(秒)、音频流信息的字典
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"视频文件不存在：{input_path}")

    # 使用ffprobe获取视频信息
    cmd = [
        'ffprobe',
        '-v', 'quiet',
        '-print_format', 'json',
        '-show_format',
        '-show_streams',
        input_path
    ]

    try:
        result = subprocess.check_output(cmd, stderr=subprocess.STDOUT).decode('utf-8')
        info = json.loads(result)

        # 获取视频时长（秒）
        duration = float(info['format']['duration'])

        # 获取音频流信息（默认取第一个音频流）
        audio_bitrate = 128000  # 默认音频码率128k bit/s（预留音频空间）
        for stream in info['streams']:
            if stream['codec_type'] == 'audio':
                if 'bit_rate' in stream and stream['bit_rate']:
                    audio_bitrate = int(stream['bit_rate'])
                break

        return {
            'duration': duration,
            'audio_bitrate': audio_bitrate
        }
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"获取视频信息失败：{e.output.decode('utf-8')}")
    except KeyError as e:
        raise RuntimeError(f"解析视频信息失败，缺少关键字段：{e}")


def calculate_target_bitrate(video_info, target_size_mb=300):
    """
    根据目标大小计算视频码率
    公式：文件大小(字节) = (视频码率 + 音频码率) * 时长(秒) / 8
    推导：视频码率 = (目标大小*1024*1024*8 / 时长) - 音频码率
    :param video_info: 视频信息字典（包含duration和audio_bitrate）
    :param target_size_mb: 目标大小（MB）
    :return: 目标视频码率（bit/s）
    """
    target_size_bytes = target_size_mb * 1024 * 1024  # 转换为字节
    total_bitrate = (target_size_bytes * 8) / video_info['duration']  # 总码率（视频+音频）
    video_bitrate = total_bitrate - video_info['audio_bitrate']

    # 确保视频码率不为负数（如果目标大小过小，至少保留基础码率）
    if video_bitrate < 100000:  # 最小视频码率100k bit/s
        video_bitrate = 100000
        print(f"警告：目标大小过小，视频码率已设为最小值100k bit/s，最终文件可能超过{target_size_mb}MB")

    return int(video_bitrate)


def compress_video(input_path, output_path, target_size_mb=300):
    """
    压缩视频到指定大小
    :param input_path: 输入视频路径
    :param output_path: 输出视频路径
    :param target_size_mb: 目标大小（MB），默认300
    """
    # 1. 获取视频信息
    video_info = get_video_info(input_path)
    duration = video_info['duration']
    audio_bitrate = video_info['audio_bitrate']

    # 2. 计算目标视频码率
    video_bitrate = calculate_target_bitrate(video_info, target_size_mb)
    print(f"视频时长：{duration:.2f}秒")
    print(f"音频码率：{audio_bitrate / 1000:.0f}k bit/s")
    print(f"目标视频码率：{video_bitrate / 1000:.0f}k bit/s")

    # 3. 执行压缩命令（使用libx264编码器，平衡压缩率和兼容性）
    cmd = [
        'ffmpeg',
        '-i', input_path,  # 输入文件
        '-c:v', 'libx264',  # 视频编码器（H.264）
        '-b:v', f'{video_bitrate}',  # 视频码率
        '-preset', 'medium',  # 压缩预设（fast=快但压缩率低，slow=慢但压缩率高）
        '-crf', '23',  # 质量控制（18-28为宜，值越小质量越高）
        '-c:a', 'aac',  # 音频编码器
        '-b:a', f'{audio_bitrate}',  # 音频码率
        '-y',  # 覆盖输出文件
        output_path  # 输出文件
    ]

    try:
        print("开始压缩视频...")
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # 检查输出文件大小
        output_size = os.path.getsize(output_path) / (1024 * 1024)
        print(f"压缩完成！")
        print(f"输出文件路径：{output_path}")
        print(f"实际文件大小：{output_size:.2f}MB")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"压缩失败：{e.stderr.decode('utf-8')}")


# 示例使用
if __name__ == "__main__":
    # 请修改以下路径和参数
    INPUT_VIDEO = "E:\System\download\IMG_1300.MOV"  # 你的源视频路径
    OUTPUT_VIDEO = "output_300mb.mp4"  # 压缩后的视频路径
    TARGET_SIZE = 300  # 目标大小（MB）

    try:
        compress_video(INPUT_VIDEO, OUTPUT_VIDEO, TARGET_SIZE)
    except Exception as e:
        print(f"错误：{e}")