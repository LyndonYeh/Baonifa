import subprocess
import sys
import os

print("安裝必要套件...")
req = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")
subprocess.run([sys.executable, "-m", "pip", "install", "-r", req], check=True)
print("\n安裝完成！")
input("按 Enter 關閉...")
