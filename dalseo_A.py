from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import pandas as pd
import time
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

SAVE_PATH = r"G:\내 드라이브\배민 자동화\달서A_실적.xlsx"
HTML_PATH = r"G:\내 드라이브\배민 자동화\달서A_dashboard.html"
REFRESH_MINUTES = 1

SET_RULES = {
    "월": [21, 20, 30, 29],
    "화": [21, 20, 30, 29],
    "수": [21, 20, 30, 29],
    "목": [21, 20, 30, 29],
    "금": [24, 21, 32, 33],
    "토": [31, 22, 36, 31],
    "일": [33, 22, 35, 30],
}

DALSEO_TEAM = [
    "김기헌", "김기현", "김민서", "김민승", "김민우", "김범준", "김병국",
    "김병철", "김영빈", "김영철", "김용우", "김우진", "김탁기", "김태광",
    "김형민", "김혜성", "나미영", "남수현", "남승훈", "문승수", "박기홍",
    "박지원", "배재현", "신진관", "신호준", "양혜진", "여재환", "윤창근",
    "이윤석", "이주철", "임상완", "임선미", "정성훈", "정승덕", "정영훈",
    "정주현", "정판호", "조대영", "황호용", "이상민"
]


def today_korean_weekday():
    now = datetime.now()
    if now.hour < 3 or (now.hour == 3 and now.minute < 10):
        now = now - timedelta(days=1)
    return ["월", "화", "수", "목", "금", "토", "일"][now.weekday()]


def click_page(driver, page_number):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)

    script = """
    const target = arguments[0];
    const buttons = Array.from(document.querySelectorAll('button'));
    const btn = buttons.find(b => b.innerText.trim() === target);
    if (btn) {
        btn.click();
        return true;
    }
    return false;
    """

    clicked = driver.execute_script(script, str(page_number))
    time.sleep(4)
    return clicked


def safe_rate(complete, reject, cancel, delivery_cancel):
    total = complete + reject + cancel + delivery_cancel
    if total <= 0:
        return 0
    return round((complete / total) * 100, 1)


def make_dashboard_html(df):
    weekday = today_korean_weekday()
    morning_set, afternoon_set, dinner_set, night_set = SET_RULES[weekday]

    dalseo_df = df[df["이름"].isin(DALSEO_TEAM)].copy()
    sonic_df = df[~df["이름"].isin(DALSEO_TEAM)].copy()

    total_complete = int(df["총완료"].sum())
    total_reject = int(df["거절"].sum())
    total_dispatch_cancel = int(df["취소"].sum())
    total_delivery_cancel = int(df["배달취소"].sum())

    total_requests = total_complete + total_reject + total_dispatch_cancel + total_delivery_cancel
    daily_accept_rate = safe_rate(total_complete, total_reject, total_dispatch_cancel, total_delivery_cancel)
    weekly_accept_rate = daily_accept_rate
    available_rejects = int(max((total_complete / 0.8) - total_requests, 0)) if total_complete else 0

    total_running = len(df[df["운행상태"].astype(str).str.contains("운행중", na=False)])
    dalseo_running = len(dalseo_df[dalseo_df["운행상태"].astype(str).str.contains("운행중", na=False)])
    sonic_running = total_running - dalseo_running

    updated = datetime.now().strftime("%Y.%m.%d %H:%M 기준")

    def width(value, target):
        if target <= 0:
            return 0
        return min(round((value / target) * 100, 1), 100)

    def peak_card(title, col, set_target):
        sonic_value = int(sonic_df[col].sum())
        dalseo_value = int(dalseo_df[col].sum())

        sonic_target = set_target * 7
        dalseo_target = set_target
        total_value = sonic_value + dalseo_value
        total_target = set_target * 8

        return f"""
        <div class="peak-card">
            <div class="peak-title">
                <div class="mini-logo">S</div>
                <span>{title} ({total_value}/{total_target})</span>
            </div>

            <div class="bar-row">
                <div class="bar-label">소닉</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{width(sonic_value, sonic_target)}%"></div>
                    <div class="bar-text">{sonic_value}/{sonic_target}</div>
                </div>
            </div>

            <div class="bar-row">
                <div class="bar-label">달서</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{width(dalseo_value, dalseo_target)}%"></div>
                    <div class="bar-text">{dalseo_value}/{dalseo_target}</div>
                </div>
            </div>
        </div>
        """

    peak_cards = "".join([
        peak_card("오전피크", "오전피크", morning_set),
        peak_card("오후논피크", "오후논피크", afternoon_set),
        peak_card("저녁피크", "저녁피크", dinner_set),
        peak_card("심야논피크", "심야논피크", night_set),
    ])

    html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="refresh" content="60">
<title>슈퍼소닉 달서A</title>
<style>
* {{ box-sizing: border-box; }}
body {{
    margin:0;
    padding:24px 14px 40px;
    background:#ffffff;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans KR",sans-serif;
    color:#111;
}}
.wrap {{ max-width:980px; margin:0 auto; }}
.logo {{
    width:70px; height:70px; margin:0 auto 12px; border-radius:50%;
    background:#ff1630; color:white; display:flex; align-items:center; justify-content:center;
    font-size:42px; font-weight:900;
}}
.title {{ text-align:center; font-size:42px; font-weight:900; margin:0; }}
.updated {{ margin-top:10px; text-align:center; color:#666; font-size:16px; font-weight:700; margin-bottom:26px; }}
.summary {{
    display:grid; grid-template-columns:repeat(2,1fr);
    border:3px solid #ff1630; border-radius:18px; overflow:hidden; margin-bottom:28px;
}}
.summary-item {{
    min-height:108px; display:flex; flex-direction:column; align-items:center; justify-content:center;
    border-right:2px solid #ff1630; border-bottom:2px solid #ff1630;
}}
.summary-item:nth-child(2n) {{ border-right:none; }}
.summary-item:nth-last-child(-n+2) {{ border-bottom:none; }}
.summary-label {{ font-size:20px; font-weight:900; margin-bottom:8px; }}
.summary-value {{ font-size:32px; font-weight:900; }}
.red {{ color:#e60012; }}
.blue {{ color:#2563eb; }}
.green {{ color:#079b24; }}
.yellow {{ color:#d97706; }}
.peaks {{ display:grid; grid-template-columns:1fr; gap:18px; }}
.peak-card {{ border:3px solid #ff1630; border-radius:22px; padding:22px; }}
.peak-title {{ display:flex; align-items:center; gap:12px; font-size:24px; font-weight:900; margin-bottom:22px; }}
.mini-logo {{
    width:38px; height:38px; border-radius:50%; background:#ff1630; color:white;
    display:flex; align-items:center; justify-content:center; font-size:22px; font-weight:900;
}}
.bar-row {{ display:grid; grid-template-columns:70px 1fr; gap:12px; align-items:center; margin-bottom:16px; }}
.bar-label {{ font-size:22px; font-weight:900; }}
.bar-wrap {{ position:relative; height:28px; background:#ffd1d8; border-radius:999px; overflow:hidden; }}
.bar-fill {{ height:100%; background:linear-gradient(90deg,#ff3b50,#ff1630); }}
.bar-text {{
    position:absolute; inset:0; display:flex; align-items:center; justify-content:center;
    font-size:18px; font-weight:900;
}}
.footer {{ margin-top:24px; text-align:center; color:#666; font-size:18px; font-weight:700; }}
</style>
</head>
<body>
<div class="wrap">
    <div class="logo">S</div>
    <h1 class="title">슈퍼소닉 달서A</h1>
    <div class="updated">🕘 {updated}</div>

    <div class="summary">
        <div class="summary-item"><div class="summary-label">총완료</div><div class="summary-value red">{total_complete:,}</div></div>
        <div class="summary-item"><div class="summary-label">총거절</div><div class="summary-value red">{total_reject:,}</div></div>
        <div class="summary-item"><div class="summary-label">배차취소</div><div class="summary-value red">{total_dispatch_cancel:,}</div></div>
        <div class="summary-item"><div class="summary-label">배달취소</div><div class="summary-value red">{total_delivery_cancel:,}</div></div>
        <div class="summary-item"><div class="summary-label">주간수락률</div><div class="summary-value yellow">{weekly_accept_rate}%</div></div>
        <div class="summary-item"><div class="summary-label">당일수락률</div><div class="summary-value blue">{daily_accept_rate}%</div></div>
        <div class="summary-item"><div class="summary-label">거절가능</div><div class="summary-value yellow">{available_rejects}개</div></div>
        <div class="summary-item"><div class="summary-label">전체접속</div><div class="summary-value green">{total_running}명</div></div>
        <div class="summary-item"><div class="summary-label">소닉팀</div><div class="summary-value green">{sonic_running}명</div></div>
        <div class="summary-item"><div class="summary-label">달서팀</div><div class="summary-value green">{dalseo_running}명</div></div>
    </div>

    <div class="peaks">
        {peak_cards}
    </div>

    <div class="footer">↻ 1분마다 자동 갱신 중...</div>
</div>
</body>
</html>
"""

    Path(HTML_PATH).write_text(html, encoding="utf-8")
    Path("달서A_dashboard.html").write_text(html, encoding="utf-8")
    print("달서A 대시보드 생성 완료")


def upload_to_github():
    print("GitHub 자동 업로드 시작")
    subprocess.run("python upload_github.py", shell=True)
    print("GitHub 자동 업로드 요청 완료")


def save_excel(results):
    if not results:
        print("저장할 데이터가 없습니다.")
        return

    df = pd.DataFrame(results).drop_duplicates()

    df["정렬순서"] = df["운행상태"].apply(lambda x: 0 if "운행중" in str(x) else 1)
    df = df.sort_values(by=["정렬순서", "총완료", "이름"], ascending=[True, False, True])
    df = df.drop(columns=["정렬순서"])

    total_row = {
        "이름": "합계",
        "운행상태": "",
        "총완료": df["총완료"].sum(),
        "거절": df["거절"].sum(),
        "취소": df["취소"].sum(),
        "배달취소": df["배달취소"].sum(),
        "오전피크": df["오전피크"].sum(),
        "오후논피크": df["오후논피크"].sum(),
        "저녁피크": df["저녁피크"].sum(),
        "심야논피크": df["심야논피크"].sum(),
        "아이디": "",
    }

    excel_df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    excel_df.to_excel(SAVE_PATH, index=False)

    make_dashboard_html(df)
    upload_to_github()

    print("\n달서A 엑셀 저장 완료")
    print(SAVE_PATH)
    print("HTML + GitHub 업로드 완료")


def collect_current_pages(driver):
    results = []

    for page in [1, 2, 3, 4, 5, 6, 7, 8, 9]:
        print(f"\n--- {page}페이지 확인 중 ---")

        clicked = click_page(driver, page)

        if not clicked:
            print(f"{page}페이지 버튼을 못 찾았습니다.")

        rows = driver.find_elements(By.TAG_NAME, "tr")

        for row in rows:
            text = row.text.strip()

            if not text:
                continue

            data = text.split("\n")

            try:
                if len(data) >= 11 and data[0] != "합계":
                    results.append({
                        "이름": data[0],
                        "운행상태": data[1],
                        "총완료": int(data[3]),
                        "거절": int(data[4]),
                        "취소": int(data[5]),
                        "배달취소": int(data[6]),
                        "오전피크": int(data[7]),
                        "오후논피크": int(data[8]),
                        "저녁피크": int(data[9]),
                        "심야논피크": int(data[10]),
                        "아이디": data[11] if len(data) > 11 else "",
                    })

                    print(f"저장: {data[0]}")

            except:
                pass

    return results


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://deliverycenter.baemin.com")

print("달서A 권역으로 로그인 후 기사목록 화면까지 이동하세요.")
input("기사목록 화면이면 엔터 누르세요...")

while True:
    print("\n화면 새로고침 중...")
    driver.refresh()
    time.sleep(6)

    results = collect_current_pages(driver)
    save_excel(results)

    print(f"\n{REFRESH_MINUTES}분 뒤 다시 갱신합니다.")
    print("종료하려면 CMD 창에서 Ctrl + C 누르세요.")

    time.sleep(REFRESH_MINUTES * 60)
