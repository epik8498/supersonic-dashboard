from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import time
from urllib.parse import quote

SMS_BODY = "공지사항 확인 부탁드립니다."
OUTPUT_PATH = r"G:\내 드라이브\배민 자동화\중구A_문자링크.txt"

my_riders = [
    "강우기", "강준호", "곽동욱", "권예준", "김도훈", "김동현", "김민", "김민승",
    "김보규", "김상일", "김용상", "김윤재", "김일수", "김지원", "김희경",
    "나원국", "노준우", "박경진", "박동혁", "박시우", "박재현", "박정권",
    "박진석", "박태우", "서정원", "손동국", "신정훈", "이국부", "이근화",
    "이상원", "이상호", "이승욱", "이재민", "임현철", "정창규", "진우일",
    "최순", "최준수", "최준호", "황정훈", "황주호", "박경보", "김장현",
    "길강호", "천용진", "이정빈", "김종찬", "성기창","윤영훈",
]


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
    time.sleep(3)
    return clicked


def clean_phone(phone):
    phone = str(phone).strip()
    phone = phone.replace("-", "").replace(" ", "")
    return phone


def make_sms_link(phone_list):
    body = quote(SMS_BODY)
    numbers = ",".join(phone_list)
    return f"sms:{numbers}?body={body}"


def collect_phones(driver):
    phones = []

    for page in [1, 2, 3, 4]:
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
                if len(data) >= 3 and data[0] != "합계":
                    name = data[0].strip()

                    if name not in my_riders:
                        continue

                    phone = clean_phone(data[2])

                    if phone.startswith("010") and len(phone) >= 10:
                        phones.append(phone)
                        print(f"번호 저장: {name} / {phone}")

            except:
                pass

    phones = list(dict.fromkeys(phones))
    return phones


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://deliverycenter.baemin.com")

print("중구A 권역으로 로그인 후 기사목록 화면까지 이동하세요.")
input("기사목록 화면이면 엔터 누르세요...")

phones = collect_phones(driver)

group_a = phones[:99]
group_b = phones[99:198]

result_text = ""

result_text += "===== 중구A 문자 링크 =====\n\n"
result_text += f"총 대상자: {len(phones)}명\n"
result_text += f"A그룹: {len(group_a)}명\n"
result_text += f"B그룹: {len(group_b)}명\n\n"

if group_a:
    result_text += "중구A 문자 링크:\n"
    result_text += make_sms_link(group_a) + "\n\n"

if group_b:
    result_text += "중구A B그룹 링크:\n"
    result_text += make_sms_link(group_b) + "\n\n"

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write(result_text)

print("\n" + result_text)
print("문자 링크 파일 저장 완료")
print(OUTPUT_PATH)

input("확인했으면 엔터 누르세요.")
driver.quit()
