import subprocess
from datetime import datetime

def run(cmd):
    subprocess.run(cmd, shell=True, check=False)

now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

run("git add index.html")
run("git add 달서A_dashboard.html")
run("git add 달서B_dashboard.html")
run("git add 중구A_dashboard.html")
run(f'git commit -m "auto update {now}"')
run("git push origin main")

print("GitHub 자동 업로드 완료")
