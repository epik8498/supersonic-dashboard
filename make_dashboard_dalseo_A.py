import pandas as pd
import time
from datetime import datetime, timedelta
from pathlib import Path

EXCEL_PATH = r"G:\내 드라이브\배민 자동화\달서A_실적.xlsx"
HTML_PATH = r"G:\내 드라이브\배민 자동화\달서A_dashboard.html"

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
    "정주현", "정판호", "조대영", "황호용"
]


def business_weekday():
    now = datetime.now()
    if now.hour < 3 or (now.hour == 3 and now.minute < 10):
        now = now - timedelta(days=1)
    return ["월", "화", "수", "목", "금", "토", "일"][now.weekday()]


def safe_int(v):
    try:
        return int(v)
    except:
        return 0


def bar_width(value, target):
    if target <= 0:
        return 0
    return min(round(value / target * 100, 1), 100)


def load_data():
    df = pd.read_excel(EXCEL_PATH, header=8)
    df = df[df["이름"].notna()]

    if "합계" in df["이름"].values:
        idx = df[df["이름"] == "합계"].index[0]
        df = df.loc[:idx - 1]

    number_cols = [
        "총완료", "거절", "취소", "배달취소",
        "오전피크", "오후논피크", "저녁피크", "심야논피크"
    ]

    for col in number_cols:
        df[col] = df[col].apply(safe_int)

    return df


def calc_accept_rate(complete, reject, dispatch_cancel, delivery_cancel):
    total = complete + reject + dispatch_cancel + delivery_cancel
    if total <= 0:
        return 0
    return round((complete / total) * 100, 1)


def calc_available_rejects(complete, reject, dispatch_cancel, delivery_cancel):
    current_total = complete + reject + dispatch_cancel + delivery_cancel
    max_total_for_80 = complete / 0.8 if complete else 0
    available = int(max_total_for_80 - current_total)
    return max(available, 0)


def make_html():
    df = load_data()

    weekday = business_weekday()
    morning_set, afternoon_set, dinner_set, night_set = SET_RULES[weekday]

    dalseo_df = df[df["이름"].isin(DALSEO_TEAM)].copy()
    sonic_df = df[~df["이름"].isin(DALSEO_TEAM)].copy()

    total_complete = int(df["총완료"].sum())
    total_reject = int(df["거절"].sum())
    total_dispatch_cancel = int(df["취소"].sum())
    total_delivery_cancel = int(df["배달취소"].sum())

    daily_accept_rate = calc_accept_rate(
        total_complete,
        total_reject,
        total_dispatch_cancel,
        total_delivery_cancel
    )

    # 현재는 주간 로그 연결 전이라 당일값 기준으로 표시합니다.
    # 다음 단계에서 weekly_log 파일이 붙으면 이 값이 수~화 누적으로 바뀝니다.
    weekly_accept_rate = daily_accept_rate

    available_rejects = calc_available_rejects(
        total_complete,
        total_reject,
        total_dispatch_cancel,
        total_delivery_cancel
    )

    total_running = len(df[df["운행상태"].astype(str).str.contains("운행중", na=False)])
    dalseo_running = len(dalseo_df[dalseo_df["운행상태"].astype(str).str.contains("운행중", na=False)])
    sonic_running = total_running - dalseo_running

    updated = datetime.now().strftime("%Y.%m.%d %H:%M 기준")

    peak_map = [
        ("오전피크", "오전피크", morning_set),
        ("오후논피크", "오후논피크", afternoon_set),
        ("저녁피크", "저녁피크", dinner_set),
        ("심야논피크", "심야논피크", night_set),
    ]

    def peak_card(title, col, set_target):
        sonic_value = int(sonic_df[col].sum())
        dalseo_value = int(dalseo_df[col].sum())

        sonic_target = set_target * 7
        dalseo_target = set_target
        total_value = sonic_value + dalseo_value
        total_target = set_target * 8

        sonic_width = bar_width(sonic_value, sonic_target)
        dalseo_width = bar_width(dalseo_value, dalseo_target)

        return f"""
        <div class="peak-card">
            <div class="peak-title">
                <div class="mini-logo">S</div>
                <span>{title} ({total_value}/{total_target})</span>
            </div>

            <div class="bar-row">
                <div class="bar-label">소닉</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{sonic_width}%"></div>
                    <div class="bar-text">{sonic_value}/{sonic_target}</div>
                </div>
            </div>

            <div class="bar-row">
                <div class="bar-label">달서</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{dalseo_width}%"></div>
                    <div class="bar-text">{dalseo_value}/{dalseo_target}</div>
                </div>
            </div>
        </div>
        """

    html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="refresh" content="60">
<title>슈퍼소닉 달서A</title>

<style>
* {{
    box-sizing: border-box;
}}

body {{
    margin: 0;
    padding: 24px 14px 40px;
    background: #ffffff;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans KR", sans-serif;
    color: #111;
}}

.wrap {{
    max-width: 980px;
    margin: 0 auto;
}}

.logo {{
    width: 70px;
    height: 70px;
    margin: 0 auto 12px;
    border-radius: 50%;
    background: #ff1630;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 42px;
    font-weight: 900;
}}

.title {{
    text-align: center;
    font-size: 42px;
    font-weight: 900;
    margin: 0;
}}

.updated {{
    margin-top: 10px;
    text-align: center;
    color: #666;
    font-size: 18px;
    font-weight: 700;
}}

.summary {{
    margin: 32px auto 34px;
    border: 3px solid #ff1630;
    border-radius: 18px;
    overflow: hidden;
    display: grid;
    grid-template-columns: repeat(2, 1fr);
}}

.summary-item {{
    min-height: 112px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;

    border-right: 2px solid #ff1630;
    border-bottom: 2px solid #ff1630;
}}

.summary-item:nth-child(2n) {{
    border-right: none;
}}

.summary-item:nth-last-child(-n+2) {{
    border-bottom: none;
}}

.summary-label {{
    font-size: 23px;
    font-weight: 900;
    margin-bottom: 9px;
}}

.summary-value {{
    font-size: 36px;
    font-weight: 900;
}}

.red {{ color: #e60012; }}
.blue {{ color: #2563eb; }}
.green {{ color: #079b24; }}
.yellow {{ color: #d97706; }}

.peaks {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 18px;
}}

.peak-card {{
    border: 3px solid #ff1630;
    border-radius: 22px;
    padding: 22px 24px;
    min-height: 220px;
}}

.peak-title {{
    display: flex;
    align-items: center;
    gap: 12px;
    font-size: 26px;
    font-weight: 900;
    margin-bottom: 24px;
}}

.mini-logo {{
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background: #ff1630;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    font-weight: 900;
}}

.bar-row {{
    display: grid;
    grid-template-columns: 74px 1fr;
    gap: 14px;
    align-items: center;
    margin-bottom: 18px;
}}

.bar-label {{
    font-size: 26px;
    font-weight: 900;
}}

.bar-wrap {{
    position: relative;
    height: 30px;
    background: #ffd1d8;
    border-radius: 999px;
    overflow: hidden;
}}

.bar-fill {{
    height: 100%;
    background: linear-gradient(90deg, #ff3b50, #ff1630);
    border-radius: 999px;
}}

.bar-text {{
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 19px;
    font-weight: 900;
}}

.footer {{
    margin-top: 28px;
    text-align: center;
    color: #777;
    font-size: 20px;
    font-weight: 800;
}}

@media (max-width:760px) {{
    .title {{
        font-size: 34px;
    }}

    .updated {{
        font-size: 15px;
    }}

    .summary-item {{
        min-height: 102px;
    }}

    .summary-label {{
        font-size: 19px;
    }}

    .summary-value {{
        font-size: 31px;
    }}

    .peaks {{
        grid-template-columns: 1fr;
    }}

    .peak-card {{
        min-height: 200px;
        padding: 20px 18px;
    }}

    .peak-title {{
        font-size: 24px;
    }}

    .bar-row {{
        grid-template-columns: 62px 1fr;
    }}

    .bar-label {{
        font-size: 24px;
    }}
}}
</style>
</head>

<body>
<div class="wrap">

    <div class="logo">S</div>
    <h1 class="title">슈퍼소닉 달서A</h1>
    <div class="updated">🕘 {updated}</div>

    <div class="summary">

        <div class="summary-item">
            <div class="summary-label">총 완료</div>
            <div class="summary-value red">{total_complete:,}</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">총 거절</div>
            <div class="summary-value red">{total_reject:,}</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">배차취소</div>
            <div class="summary-value red">{total_dispatch_cancel:,}개</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">배달취소</div>
            <div class="summary-value red">{total_delivery_cancel:,}개</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">주간수락률</div>
            <div class="summary-value yellow">{weekly_accept_rate}%</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">당일수락률</div>
            <div class="summary-value blue">{daily_accept_rate}%</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">거절가능</div>
            <div class="summary-value yellow">{available_rejects}개</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">전체접속</div>
            <div class="summary-value green">{total_running}명</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">소닉팀</div>
            <div class="summary-value green">{sonic_running}명</div>
        </div>

        <div class="summary-item">
            <div class="summary-label">달서팀</div>
            <div class="summary-value green">{dalseo_running}명</div>
        </div>

    </div>

    <div class="peaks">
        {''.join([peak_card(title, col, target) for title, col, target in peak_map])}
    </div>

    <div class="footer">↻ 1분마다 자동 갱신 중...</div>

</div>
</body>
</html>
"""

    Path(HTML_PATH).write_text(html, encoding="utf-8")

    print("달서A 대시보드 생성 완료")
    print(HTML_PATH)


if __name__ == "__main__":
    while True:
        make_html()
        print("1분 뒤 다시 갱신합니다.")
        time.sleep(60)
