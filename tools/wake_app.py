"""Streamlit 앱이 잠들어 있으면 깨우는 스크립트 (GitHub Actions 에서 실행)."""
import sys
import time

from playwright.sync_api import sync_playwright

URL = "https://biopharm-precost.streamlit.app/"
WAKE_TEXT = "get this app back up"
READY_TEXTS = ["사전원가 자동입력", "아이디"]
TIMEOUT_SEC = 300


def page_has(page, text):
    for frame in page.frames:
        try:
            if frame.get_by_text(text, exact=False).count() > 0:
                return True
        except Exception:
            pass
    return False


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=120_000)
        time.sleep(15)

        woke = False
        for frame in page.frames:
            btn = frame.get_by_role("button", name=WAKE_TEXT)
            if btn.count() > 0:
                print("앱이 잠들어 있어 깨우기 버튼을 누릅니다.")
                btn.first.click()
                woke = True
                break
        if not woke:
            print("잠자기 화면이 아닙니다. 정상 화면이 뜨는지 확인합니다.")

        start = time.time()
        while time.time() - start < TIMEOUT_SEC:
            if any(page_has(page, t) for t in READY_TEXTS):
                print(f"정상: 로그인 화면 확인 ({int(time.time() - start)}초)")
                browser.close()
                return 0
            time.sleep(10)

        page.screenshot(path="wake_fail.png", full_page=True)
        print("실패: 제한 시간 안에 로그인 화면이 뜨지 않았습니다.")
        browser.close()
        return 1


if __name__ == "__main__":
    sys.exit(main())
