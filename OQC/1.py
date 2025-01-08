from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
import os

# 生成 32 字节（256 位）密钥和 16 字节（128 位）IV
key = os.urandom(32)
iv = os.urandom(16)

# 创建 AES-256 加密器，CBC 模式
cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())

# 数据填充（PKCS7）
padder = padding.PKCS7(128).padder()
plaintext = b'This is a secret message.'
padded_data = padder.update(plaintext) + padder.finalize()

# 加密
encryptor = cipher.encryptor()
ciphertext = encryptor.update(padded_data) + encryptor.finalize()

# 打印加密结果
print(f"Key: {key.hex()}")
print(f"IV: {iv.hex()}")
print(f"Ciphertext: {ciphertext.hex()}")
