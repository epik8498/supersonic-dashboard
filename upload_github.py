import subprocess
from datetime import datetime

msg = datetime.now().strftime("auto update %Y-%m-%d %H:%M:%S")

commands = [
    "git add .",
    f'git commit -m "{msg}"',
    "git pull --rebase origin main",
    "git push origin main"
]

for cmd in commands:
    print(f"\n실행: {cmd}")
    result = subprocess.run(cmd, shell=True)

print("\nGitHub 자동 업로드 완료")