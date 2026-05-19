from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import pandas as pd
import time
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

SAVE_PATH = r"G:\내 드라이브\배민 자동화\달서B_실적.xlsx"
HTML_PATH = r"G:\내 드라이브\배민 자동화\달서B_dashboard.html"
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

TEAM_SETS = {
    "넘버": 6,
    "마음": 5,
    "슈퍼": 2,
}

TEAM_MAP_RAW = {
    "유호성":"넘버","박세창":"넘버","강명원":"넘버","김수진":"넘버","배서후":"넘버","김요한":"넘버",
    "김정근":"넘버","남승호":"넘버","이헌재":"넘버","이윤재":"넘버","정수영":"넘버","장정석":"넘버",
    "최영진":"넘버","임현석":"넘버","임승범":"넘버","이태훈":"넘버","이철우":"넘버","이재헌":"넘버",
    "이은성":"넘버","이영희":"넘버","이선노":"넘버","이동석":"넘버","우효상":"넘버","서강원":"넘버",
    "한동훈":"넘버","마경민":"넘버","노재권":"넘버","남윤정":"넘버","남동욱":"넘버","김현준":"넘버",
    "김태하":"넘버","김종희":"넘버","김용운":"넘버","김영천":"넘버","김병수":"넘버","김명한":"넘버",
    "김동국":"넘버","권오현":"넘버","황홍섭":"넘버","강지은":"넘버","최윤호":"넘버","신명섭":"넘버",
    "윤민석":"넘버","김애선":"넘버","이대겸":"넘버","김대운":"넘버",

    "길강호":"마음","김영우":"마음","강지우":"마음","이승훈":"마음","박성립":"마음","이영민":"마음",
    "손성곤":"마음","구상훈":"마음","박한울":"마음","신가희":"마음","박연호":"마음","김형택":"마음",
    "김낙훈":"마음","권영남":"마음","이진복":"마음","김석원":"마음","길태빈":"마음","김창범":"마음",
    "박광용":"마음","성영길":"마음","박원희":"마음","임용우":"마음","최영우":"마음","이전필":"마음",
    "이재현":"마음","이강현":"마음","김대환":"마음","여세동":"마음","신정학":"마음","임지훈":"마음",
    "장민서":"마음","임종헌":"마음","윤동근":"마음","도수현":"마음","김동현":"마음","정동진":"마음",
    "정동수":"마음","전현":"마음","전하경":"마음","전승옥":"마음","전대명":"마음","장예환":"마음",
    "장대웅":"마음","임재백":"마음","이진욱":"마음","이진승":"마음","이승준":"마음","이경태":"마음",
    "최현준":"마음","안호식":"마음","신원준":"마음","서봉용":"마음","박호일":"마음","박성우":"마음",
    "도인환":"마음","노지훈":"마음","김현진":"마음","김지성":"마음","김재훈":"마음","황유경":"마음",
    "김성현":"마음","김서현":"마음","김효겸":"마음","송인섭":"마음","김종서":"마음","김종호":"마음",
    "남재화":"마음","박남아":"마음","구용태":"마음","한대성":"마음","윤정원":"마음","손지수":"마음",
    "김숙자":"마음","김현숙":"마음","최종현":"마음","김인수":"마음","김임식":"마음","신인호":"마음",
    "구자돈":"마음","백병준":"마음","차무길":"마음","차성원":"마음","박지홍":"마음","이예준":"마음",
    "위석훈":"마음","피우덕":"마음","소귀숙":"마음","피우정":"마음","백창열":"마음","하태수":"마음",
    "명제규":"마음","한희숙":"마음","김동욱":"마음","김도형":"마음",

    "최경민":"슈퍼","윤규범":"슈퍼","신성욱":"슈퍼","박무성":"슈퍼","송특근":"슈퍼","정우혁":"슈퍼",
    "김경섭":"슈퍼","장근영":"슈퍼","조윤환":"슈퍼","조승래":"슈퍼","정기정":"슈퍼","정규태":"슈퍼",
    "장재근":"슈퍼","최지나":"슈퍼","이종필":"슈퍼","이정민":"슈퍼","이재상":"슈퍼","이재관":"슈퍼",
    "윤철훈":"슈퍼","유영엽":"슈퍼","엄정철":"슈퍼","심재득":"슈퍼","신진학":"슈퍼","배준호":"슈퍼",
    "박정민":"슈퍼","김주동":"슈퍼","김재현":"슈퍼","김상엽":"슈퍼","김동규":"슈퍼","권휘재":"슈퍼",
    "강현구":"슈퍼","최지웅":"슈퍼","김종찬":"슈퍼","이상엽":"슈퍼",
}

TEAM_MAP = dict(TEAM_MAP_RAW)


def today_korean_weekday():
    now = datetime.now()
    if now.hour < 3 or (now.hour == 3 and now.minute < 10):
        now = now - timedelta(days=1)
    return ["월", "화", "수", "목", "금", "토", "일"][now.weekday()]


def get_team(name):
    return TEAM_MAP.get(str(name).strip(), "미분류")


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
    return round((complete / total) * 100, 1) if total else 0


def make_dashboard_html(df):
    weekday = today_korean_weekday()
    morning_set, afternoon_set, dinner_set, night_set = SET_RULES[weekday]

    df["팀"] = df["이름"].apply(get_team)

    number_df = df[df["팀"] == "넘버"].copy()
    maeum_df = df[df["팀"] == "마음"].copy()
    super_df = df[df["팀"] == "슈퍼"].copy()
    unknown_df = df[df["팀"] == "미분류"].copy()

    total_complete = int(df["총완료"].sum())
    total_reject = int(df["거절"].sum())
    total_dispatch_cancel = int(df["취소"].sum())
    total_delivery_cancel = int(df["배달취소"].sum())

    total_requests = total_complete + total_reject + total_dispatch_cancel + total_delivery_cancel
    daily_accept_rate = safe_rate(total_complete, total_reject, total_dispatch_cancel, total_delivery_cancel)
    weekly_accept_rate = daily_accept_rate
    available_rejects = int(max((total_complete / 0.8) - total_requests, 0)) if total_complete else 0

    total_running = len(df[df["운행상태"].astype(str).str.contains("운행중", na=False)])
    number_running = len(number_df[number_df["운행상태"].astype(str).str.contains("운행중", na=False)])
    maeum_running = len(maeum_df[maeum_df["운행상태"].astype(str).str.contains("운행중", na=False)])
    super_running = len(super_df[super_df["운행상태"].astype(str).str.contains("운행중", na=False)])

    updated = datetime.now().strftime("%Y.%m.%d %H:%M 기준")

    def width(value, target):
        return min(round((value / target) * 100, 1), 100) if target else 0

    def peak_card(title, col, set_target):
        number_value = int(number_df[col].sum())
        maeum_value = int(maeum_df[col].sum())
        super_value = int(super_df[col].sum())
        unknown_value = int(unknown_df[col].sum())

        number_target = set_target * TEAM_SETS["넘버"]
        maeum_target = set_target * TEAM_SETS["마음"]
        super_target = set_target * TEAM_SETS["슈퍼"]

        total_value = number_value + maeum_value + super_value + unknown_value
        total_target = set_target * 13

        unknown_html = ""
        if unknown_value > 0:
            unknown_html = f"""
            <div class="bar-row">
                <div class="bar-label">미분류</div>
                <div class="bar-wrap">
                    <div class="bar-fill gray-fill" style="width:100%"></div>
                    <div class="bar-text">{unknown_value}/-</div>
                </div>
            </div>
            """

        return f"""
        <div class="peak-card">
            <div class="peak-title">
                <div class="wing">S</div>
                <span>{title} ({total_value}/{total_target})</span>
            </div>

            <div class="bar-row">
                <div class="bar-label">넘버</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{width(number_value, number_target)}%"></div>
                    <div class="bar-text">{number_value}/{number_target}</div>
                </div>
            </div>

            <div class="bar-row">
                <div class="bar-label">마음</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{width(maeum_value, maeum_target)}%"></div>
                    <div class="bar-text">{maeum_value}/{maeum_target}</div>
                </div>
            </div>

            <div class="bar-row">
                <div class="bar-label">슈퍼</div>
                <div class="bar-wrap">
                    <div class="bar-fill" style="width:{width(super_value, super_target)}%"></div>
                    <div class="bar-text">{super_value}/{super_target}</div>
                </div>
            </div>

            {unknown_html}
        </div>
        """

    peak_cards = "".join([
        peak_card("오전피크", "오전피크", morning_set),
        peak_card("오후논피크", "오후논피크", afternoon_set),
        peak_card("저녁피크", "저녁피크", dinner_set),
        peak_card("심야논피크", "심야논피크", night_set),
    ])

    rider_cards = ""

    for _, row in df.iterrows():
        name = row["이름"]
        team = row["팀"]
        status = str(row["운행상태"])
        is_online = "운행중" in status

        card_class = "rider-card online" if is_online else "rider-card offline"
        badge_class = "badge online-badge" if is_online else "badge offline-badge"
        badge_text = "접속중" if is_online else "오프라인"
        online_sort = 0 if is_online else 1

        rider_cards += f"""
        <div class="{card_class}"
            data-team="{team}"
            data-online="{online_sort}"
            data-name="{name}"
            data-complete="{int(row['총완료'])}">

            <div class="rider-top">
                <div class="rider-name">{name}</div>
                <div class="{badge_class}">{badge_text}</div>
            </div>

            <div class="rider-sub">{team} | {badge_text}</div>

            <div class="rider-stats">
                <div><span>완료</span><b class="blue">{int(row['총완료'])}</b></div>
                <div><span>거절</span><b class="red">{int(row['거절'])}</b></div>
                <div><span>오전</span><b>{int(row['오전피크'])}</b></div>
                <div><span>오후</span><b>{int(row['오후논피크'])}</b></div>
                <div><span>저녁</span><b>{int(row['저녁피크'])}</b></div>
                <div><span>심야</span><b>{int(row['심야논피크'])}</b></div>
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
<title>슈퍼소닉 달서B</title>

<style>
* {{ box-sizing:border-box; }}

body {{
    margin:0;
    background:#ffffff;
    color:#111;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans KR",sans-serif;
}}

.topbar {{
    position:sticky;
    top:0;
    z-index:10;
    background:#fff;
    border-bottom:1px solid #eee;
    display:flex;
    align-items:center;
    justify-content:space-between;
    padding:14px 34px;
}}

.top-left {{
    display:flex;
    align-items:center;
    gap:14px;
    font-size:26px;
    font-weight:900;
}}

.mark {{
    width:48px;
    height:48px;
    border-radius:50%;
    background:#ff1630;
    color:#fff;
    display:flex;
    align-items:center;
    justify-content:center;
    font-weight:900;
    font-size:27px;
    position:relative;
}}

.mark:before,
.mark:after {{
    content:"";
    position:absolute;
    width:24px;
    height:10px;
    border-top:5px solid #ff1630;
    top:16px;
}}

.mark:before {{ left:-23px; transform:rotate(24deg); }}
.mark:after {{ right:-23px; transform:rotate(-24deg); }}

.refresh {{
    color:#555;
    font-size:14px;
    line-height:1.35;
    font-weight:700;
    text-align:right;
}}

.wrap {{
    max-width:1180px;
    margin:0 auto;
    padding:34px 18px 50px;
}}

.hero-logo {{
    width:72px;
    height:72px;
    margin:0 auto 12px;
    border-radius:50%;
    background:#ff1630;
    color:#fff;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:42px;
    font-weight:900;
    position:relative;
}}

.hero-logo:before,
.hero-logo:after {{
    content:"";
    position:absolute;
    width:38px;
    height:15px;
    border-top:7px solid #ff1630;
    top:23px;
}}

.hero-logo:before {{ left:-36px; transform:rotate(24deg); }}
.hero-logo:after {{ right:-36px; transform:rotate(-24deg); }}

.title {{
    text-align:center;
    font-size:42px;
    font-weight:900;
    margin:0 0 8px;
}}

.updated {{
    text-align:center;
    color:#666;
    font-size:14px;
    margin-bottom:24px;
}}

.summary {{
    max-width:720px;
    margin:0 auto 32px;
    border:3px solid #ff1630;
    border-radius:14px;
    overflow:hidden;
    display:grid;
    grid-template-columns:repeat(2,1fr);
}}

.summary-item {{
    min-height:74px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    padding:0 26px;
    border-right:2px solid #ff1630;
    border-bottom:1.5px solid #ff1630;
    font-size:19px;
    font-weight:800;
}}

.summary-item:nth-child(2n) {{ border-right:none; }}
.summary-item:nth-last-child(-n+2) {{ border-bottom:none; }}

.summary-value {{
    font-size:23px;
    font-weight:900;
}}

.red {{ color:#e60012; }}
.blue {{ color:#1455ff; }}
.green {{ color:#05a832; }}
.yellow {{ color:#d58900; }}

.peaks {{
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:20px 28px;
    margin-bottom:24px;
}}

.peak-card {{
    border:3px solid #ff1630;
    border-radius:24px;
    padding:22px 34px 28px;
}}

.peak-title {{
    display:flex;
    align-items:center;
    gap:16px;
    font-size:28px;
    font-weight:900;
    margin-bottom:30px;
}}

.wing {{
    width:34px;
    height:34px;
    border-radius:50%;
    background:#ff1630;
    color:white;
    display:flex;
    align-items:center;
    justify-content:center;
    font-weight:900;
    font-size:20px;
}}

.bar-row {{
    display:grid;
    grid-template-columns:88px 1fr;
    align-items:center;
    gap:16px;
    margin-bottom:16px;
}}

.bar-label {{
    font-size:28px;
    font-weight:900;
}}

.bar-wrap {{
    position:relative;
    height:30px;
    background:#ffd1d8;
    border-radius:999px;
    overflow:hidden;
}}

.bar-fill {{
    height:100%;
    background:linear-gradient(90deg,#ff3b50,#ff1630);
    border-radius:999px;
}}

.gray-fill {{
    background:#777;
}}

.bar-text {{
    position:absolute;
    inset:0;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:20px;
    font-weight:900;
}}

.filter-area {{
    border-top:1px solid #ddd;
    padding-top:20px;
    margin-top:16px;
}}

.control-row {{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
    margin-bottom:18px;
    flex-wrap:wrap;
}}

.tabs {{
    display:flex;
    flex-wrap:wrap;
    gap:10px;
}}

.tab {{
    border:none;
    border-radius:12px;
    padding:14px 24px;
    font-size:20px;
    font-weight:900;
    background:#eee;
    cursor:pointer;
}}

.tab.active {{
    background:#08b23b;
    color:white;
}}

.sort-box {{
    display:flex;
    align-items:center;
    gap:10px;
    font-weight:900;
    font-size:18px;
}}

.sort-select {{
    padding:13px 16px;
    border:2px solid #111;
    border-radius:12px;
    font-size:18px;
    font-weight:900;
    background:#fff;
}}

.riders {{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:18px;
}}

.rider-card {{
    border:2.5px solid #555;
    border-radius:16px;
    padding:18px;
    background:white;
}}

.rider-card.online {{ border-color:#08c747; }}
.rider-card.offline {{ border-color:#444; opacity:0.82; }}

.rider-top {{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:8px;
}}

.rider-name {{
    font-size:24px;
    font-weight:900;
}}

.badge {{
    border-radius:999px;
    padding:6px 12px;
    color:white;
    font-size:15px;
    font-weight:900;
    white-space:nowrap;
}}

.online-badge {{ background:#08b23b; }}
.offline-badge {{ background:#555; }}

.rider-sub {{
    margin-top:8px;
    color:#666;
    font-size:15px;
    font-weight:700;
}}

.rider-stats {{
    margin-top:16px;
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:12px 6px;
}}

.rider-stats div {{
    display:flex;
    gap:6px;
    align-items:center;
    font-size:14px;
    font-weight:800;
}}

.rider-stats b {{ font-size:18px; }}

.footer {{
    margin-top:28px;
    text-align:center;
    color:#666;
    font-size:16px;
    font-weight:700;
}}

@media (max-width:900px) {{
    .topbar {{ padding:12px 16px; flex-direction:column; gap:10px; }}
    .top-left {{ font-size:22px; }}
    .refresh {{ text-align:center; }}
    .title {{ font-size:34px; }}
    .summary-item {{ min-height:74px; padding:0 16px; font-size:16px; }}
    .summary-value {{ font-size:21px; }}
    .peaks {{ grid-template-columns:1fr; }}
    .peak-card {{ padding:20px 18px; }}
    .peak-title {{ font-size:24px; }}
    .bar-row {{ grid-template-columns:64px 1fr; }}
    .bar-label {{ font-size:23px; }}
    .control-row {{ display:block; }}
    .tabs {{ margin-bottom:14px; }}
    .sort-box {{ justify-content:space-between; }}
    .sort-select {{ width:180px; }}
    .riders {{ grid-template-columns:repeat(2,1fr); gap:12px; }}
    .rider-name {{ font-size:20px; }}
    .tab {{ font-size:17px; padding:12px 16px; }}
}}

@media (max-width:520px) {{
    .riders {{ grid-template-columns:1fr; }}
}}
</style>
</head>

<body>

<div class="topbar">
    <div class="top-left">
        <div class="mark">S</div>
        <div>슈퍼소닉 달서B 대시보드</div>
    </div>
    <div class="refresh">🔄 1분마다 자동 업데이트</div>
</div>

<div class="wrap">

    <div class="hero-logo">S</div>
    <h1 class="title">슈퍼소닉 달서B</h1>
    <div class="updated">🕘 {updated}</div>

    <div class="summary">
        <div class="summary-item"><span>총완료</span><b class="summary-value red">{total_complete:,}</b></div>
        <div class="summary-item"><span>총거절</span><b class="summary-value red">{total_reject:,}</b></div>
        <div class="summary-item"><span>배차취소</span><b class="summary-value red">{total_dispatch_cancel:,}</b></div>
        <div class="summary-item"><span>배달취소</span><b class="summary-value red">{total_delivery_cancel:,}</b></div>
        <div class="summary-item"><span>주간수락률</span><b class="summary-value yellow">{weekly_accept_rate}%</b></div>
        <div class="summary-item"><span>당일수락률</span><b class="summary-value blue">{daily_accept_rate}%</b></div>
        <div class="summary-item"><span>거절가능</span><b class="summary-value yellow">{available_rejects}개</b></div>
        <div class="summary-item"><span>전체접속</span><b class="summary-value green">{total_running}명</b></div>
        <div class="summary-item"><span>넘버</span><b class="summary-value green">{number_running}명</b></div>
        <div class="summary-item"><span>마음</span><b class="summary-value green">{maeum_running}명</b></div>
        <div class="summary-item"><span>슈퍼</span><b class="summary-value green">{super_running}명</b></div>
    </div>

    <div class="peaks">{peak_cards}</div>

    <div class="filter-area">
        <div class="control-row">
            <div class="tabs">
                <button class="tab active" onclick="setFilter('전체', this)">전체 ({len(df)})</button>
                <button class="tab" onclick="setFilter('넘버', this)">넘버 ({len(number_df)})</button>
                <button class="tab" onclick="setFilter('마음', this)">마음 ({len(maeum_df)})</button>
                <button class="tab" onclick="setFilter('슈퍼', this)">슈퍼 ({len(super_df)})</button>
            </div>

            <div class="sort-box">
                <span>정렬</span>
                <select id="sortSelect" class="sort-select" onchange="applySort()">
                    <option value="online">접속중 우선</option>
                    <option value="name">가나다순</option>
                    <option value="complete">완료순</option>
                </select>
            </div>
        </div>

        <div id="riders" class="riders">{rider_cards}</div>
    </div>

    <div class="footer">실시간 데이터 · 1분마다 자동 업데이트</div>

</div>

<script>
let currentFilter = "전체";

function setFilter(value, btn) {{
    currentFilter = value;
    document.querySelectorAll(".tab").forEach(item => item.classList.remove("active"));
    btn.classList.add("active");
    applyView();
}}

function applySort() {{
    applyView();
}}

function applyView() {{
    const sortValue = document.getElementById("sortSelect").value;
    const container = document.getElementById("riders");
    const cards = Array.from(document.querySelectorAll(".rider-card"));

    cards.forEach(card => {{
        const team = card.dataset.team;
        let show = true;

        if(currentFilter === "넘버") show = team === "넘버";
        if(currentFilter === "마음") show = team === "마음";
        if(currentFilter === "슈퍼") show = team === "슈퍼";

        card.style.display = show ? "block" : "none";
    }});

    const visibleCards = cards.filter(card => card.style.display !== "none");

    visibleCards.sort((a, b) => {{
        if(sortValue === "online") {{
            const onlineA = Number(a.dataset.online);
            const onlineB = Number(b.dataset.online);
            if(onlineA !== onlineB) return onlineA - onlineB;
            return Number(b.dataset.complete) - Number(a.dataset.complete);
        }}

        if(sortValue === "name") {{
            return a.dataset.name.localeCompare(b.dataset.name, "ko");
        }}

        if(sortValue === "complete") {{
            return Number(b.dataset.complete) - Number(a.dataset.complete);
        }}

        return 0;
    }});

    visibleCards.forEach(card => container.appendChild(card));
}}

document.addEventListener("DOMContentLoaded", function() {{
    applyView();
}});
</script>

</body>
</html>
"""

    Path(HTML_PATH).write_text(html, encoding="utf-8")
    Path("달서B_dashboard.html").write_text(html, encoding="utf-8")
    print("달서B 대시보드 생성 완료")


def upload_to_github():
    print("GitHub 자동 업로드 시작")
    subprocess.run("python upload_github.py", shell=True)
    print("GitHub 자동 업로드 요청 완료")


def save_excel(results):
    if not results:
        print("저장할 데이터가 없습니다.")
        return

    df = pd.DataFrame(results).drop_duplicates()

    df["팀"] = df["이름"].apply(get_team)
    df["정렬순서"] = df["운행상태"].apply(lambda x: 0 if "운행중" in str(x) else 1)
    df = df.sort_values(by=["정렬순서", "팀", "총완료", "이름"], ascending=[True, True, False, True])
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
        "팀": "",
    }

    excel_df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    excel_df.to_excel(SAVE_PATH, index=False)

    make_dashboard_html(df)
    upload_to_github()

    print("\n달서B 엑셀 저장 완료")
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
                    name = data[0].strip()

                    results.append({
                        "이름": name,
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

                    print(f"저장: {name} / {get_team(name)}")

            except:
                pass

    return results


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://deliverycenter.baemin.com")

print("달서B 권역으로 로그인 후 기사목록 화면까지 이동하세요.")
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
