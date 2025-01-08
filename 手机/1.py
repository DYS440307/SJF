import numpy as np
import librosa
import librosa.display
import scipy.io.wavfile as wav
import matplotlib.pyplot as plt
import os

# 读取音频文件
file_path = r"\\192.168.11.24\惠州声乐\3-品质部\实验室\邓洋枢\1-实验室相关文件\7-测试信号\IEC268-5_no_filtre_Noise_with_crestfactor_2.wav"
save_path = r"F:\system\Pictures\转中文件夹"

# 检查保存路径是否存在，不存在则创建
if not os.path.exists(save_path):
    os.makedirs(save_path)

# 使用librosa加载音频
signal, sr = librosa.load(file_path, sr=None)  # sr=None保持原始采样率

# 计算音频的时长
duration = librosa.get_duration(y=signal, sr=sr)

# 计算频率范围
# 使用傅里叶变换计算频谱
fft_spectrum = np.fft.rfft(signal)
frequencies = np.fft.rfftfreq(len(signal), d=1/sr)
magnitude_spectrum = np.abs(fft_spectrum)

# 计算峰值因子 (CF)
peak_value = np.max(np.abs(signal))
rms_value = np.sqrt(np.mean(signal**2))
crest_factor = peak_value / rms_value

# 显示频谱图并保存
plt.figure(figsize=(10, 6))
plt.plot(frequencies, magnitude_spectrum)
plt.title('Frequency Spectrum')
plt.xlabel('Frequency (Hz)')
plt.ylabel('Amplitude')
plt.grid(True)

# 保存图像到指定路径
output_file = os.path.join(save_path, "frequency_spectrum.png")
plt.savefig(output_file)
plt.close()  # 关闭图形以释放内存

# 输出分析结果
print(f"音频时长: {duration:.2f}秒")
print(f"峰值因子 (Crest Factor): {crest_factor:.2f}")
print(f"频率范围: {frequencies[0]:.2f} Hz - {frequencies[-1]:.2f} Hz")
print(f"频谱图已保存至: {output_file}")
