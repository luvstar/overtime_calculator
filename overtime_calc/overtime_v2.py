import tkinter as tk
from tkinter import scrolledtext, messagebox
import pandas as pd
import time
from seleniumwire import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.chrome.options import Options
import json
import gzip
from datetime import datetime, timedelta
import os
import sys

# --- (1) 사용자 설정 ---
LOGIN_URL = "https://gw.cubox.ai/#/login?logout=Y&lang=kr"
TARGET_PAGE_URL = "https://gw.cubox.ai/#/HP/HPD0220/HPD0220"
INTERCEPT_URL_KEYWORD = "selectTab2"
JSON_DATA_LIST_KEY = "resultData"
DATE_KEY = "atDt"
START_TIME_KEY = "comeTm"
END_TIME_KEY = "leaveTm"

# --- UI 로깅 함수 ---
def log_to_ui(text_widget, message):
    log_bg_color = "#1E1E1E"
    text_color = "#EAEAEA"
    text_widget.config(state=tk.NORMAL, bg=log_bg_color, fg=text_color)
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END) 
    text_widget.config(state=tk.DISABLED, bg=log_bg_color, fg=text_color)
    window.update_idletasks() 

# --- (2) 핵심 로직: 근무 시간 계산 ---
def calculate_work_hours(json_data_list, text_widget):
    try:
        if not json_data_list:
            return "오류: JSON 데이터 목록이 비어있습니다."
        df = pd.DataFrame(json_data_list)
        def format_hhmm(hhmm_val):
            if pd.isna(hhmm_val) or hhmm_val == '':
                return None
            try:
                hhmm_str = str(int(float(hhmm_val))).zfill(4) 
            except ValueError:
                hhmm_str = str(hhmm_val)
            if len(hhmm_str) < 4:
                return None
            return f"{hhmm_str[:2]}:{hhmm_str[2:]}"
        df[START_TIME_KEY] = df[START_TIME_KEY].apply(format_hhmm)
        df[END_TIME_KEY] = df[END_TIME_KEY].apply(format_hhmm)
        df['날짜'] = pd.to_datetime(df[DATE_KEY]).dt.date
        df['출근시간'] = pd.to_datetime(df['날짜'].astype(str) + ' ' + df[START_TIME_KEY], errors='coerce')
        df['퇴근시간'] = pd.to_datetime(df['날짜'].astype(str) + ' ' + df[END_TIME_KEY], errors='coerce')
        original_rows = len(df)
        df = df.dropna(subset=['출근시간', '퇴근시간'])
        dropped_rows = original_rows - len(df)
        if dropped_rows > 0:
            log_to_ui(text_widget, f"알림: 출근/퇴근 시간이 없는 {dropped_rows}개의 행을 계산에서 제외했습니다.")
        if df.empty:
            return "계산할 수 있는 유효한 출퇴근 기록이 없습니다."
        df['실근무시간'] = (df['퇴근시간'] - df['출근시간']) - pd.Timedelta(hours=1)
        eight_hours = pd.Timedelta(hours=8)
        df['일일초과'] = df['실근무시간'] - eight_hours
        df['일일초과'] = df['일일초과'].apply(lambda x: max(x, pd.Timedelta(0)))
        df['날짜_dt'] = pd.to_datetime(df['날짜'])
        df['주차'] = df['날짜_dt'].dt.isocalendar().week
        weekly_summary = df.groupby('주차')['실근무시간'].sum().reset_index()
        weekly_summary.columns = ['주차', '주간총무']
        forty_hours = pd.Timedelta(hours=40)
        weekly_summary['주간초과'] = forty_hours - weekly_summary['주간총무']
        weekly_summary['주간초과'] = weekly_summary['주간초과'].apply(lambda x: max(x, pd.Timedelta(0)))
        result_text = "=== 📅 일별 초과근무 ===\n"
        if dropped_rows > 0:
            result_text += f"(참고: 유효하지 않은 {dropped_rows}개 행 제외)\n\n"
        def format_timedelta_simple(td):
            if pd.isna(td):
                return "00:00:00"
            total_seconds = td.total_seconds()
            hours = int(total_seconds // 3600)
            minutes = int((total_seconds % 3600) // 60)
            seconds = int(total_seconds % 60)
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        df['일일초과_str'] = df['일일초과'].apply(format_timedelta_simple)
        df['실근무시간_str'] = df['실근무시간'].apply(format_timedelta_simple)
        for index, row in df.iterrows():
            result_text += f"[{row['날짜']}] 일일 초과: {row['일일초과_str']} 일일 근무 시간 : {row['실근무시간_str']}\n"
        result_text += "\n\n=== 📊 주별 초과근무 ===\n\n"
        weekly_summary['주간총무_str'] = weekly_summary['주간총무'].apply(format_timedelta_simple)
        weekly_summary['주간초과_str'] = weekly_summary['주간초과'].apply(format_timedelta_simple)
        for index, row in weekly_summary.iterrows():
            result_text += f"[{row['주차']}주차] 총 근무: {row['주간총무_str']} | 남은 주간 근무 시간: {row['주간초과_str']}\n"
        return result_text
    except KeyError as e:
        return f"키 오류: {e}\n\n(1)번 사용자 설정의 JSON 키 이름(예: JSON_DATA_LIST_KEY)이\nF12 [Response] 탭의 이름과 일치하는지 확인하세요."
    except Exception as e:
        return f"계산 중 알 수 없는 오류 발생: {e}"

# --- (3) Selenium 자동화 로직 ---
def run_automation_and_calculate(user_id, user_pw, start_date, end_date, result_text_area):
    try:
        # 0. UI 업데이트
        result_text_area.config(state=tk.NORMAL)
        result_text_area.delete('1.0', tk.END) 
        result_text_area.config(state=tk.DISABLED)
        log_to_ui(result_text_area, "자동 로그인을 시작합니다...\n웹 브라우저를 백그라운드에서 실행합니다.") 

        # 헤드리스 모드(백그라운드 실행) 옵션 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080") # 가상 윈도우 크기 설정
        chrome_options.add_argument("--disable-gpu") # GPU 비활성화
        
        service = Service(ChromeDriverManager().install())
        # 옵션을 적용하여 드라이버 실행
        driver = webdriver.Chrome(service=service, options=chrome_options) 

        driver.get(LOGIN_URL)
        time.sleep(3) 
        log_to_ui(result_text_area, f"페이지 이동 시도: {LOGIN_URL}")
        
        # 3. 자동 로그인
        try:
            wait = WebDriverWait(driver, 10)
            log_to_ui(result_text_area, "ID 입력창을 기다립니다...")
            id_input = wait.until(EC.presence_of_element_located((By.ID, "reqLoginId")))
            log_to_ui(result_text_area, "ID 입력창 찾음. ID 입력...")
            id_input.send_keys(user_id)
            log_to_ui(result_text_area, "'다음' 버튼을 기다립니다...")
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '다음')]")))
            log_to_ui(result_text_area, "'다음' 버튼 클릭.")
            next_button.click()
            log_to_ui(result_text_area, "비밀번호 입력창을 기다립니다...")
            pw_input = wait.until(EC.element_to_be_clickable((By.ID, "reqLoginPw")))
            log_to_ui(result_text_area, "비밀번호 입력...")
            pw_input.send_keys(user_pw)
            log_to_ui(result_text_area, "로그인 버튼을 기다립니다...")
            login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '로그인')]")))
            log_to_ui(result_text_area, "로그인 버튼 클릭.")
            login_button.click()
        except Exception as e:
            log_to_ui(result_text_area, f"오류 발생! 현재 URL: {driver.current_url}")
            driver.quit()
            messagebox.showerror("로그인 실패", f"로그인 요소를 찾을 수 없습니다.\n오류: {e}")
            return

        log_to_ui(result_text_area, "\n로그인 성공. 데이터 페이지로 이동합니다...")

        # --- 4. 데이터 페이지 이동 및 날짜 설정 ---
        try:
            wait = WebDriverWait(driver, 10) 
            log_to_ui(result_text_area, "'근무시간현황' 메뉴/버튼을 찾습니다...")
            commute_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-name='근무시간현황']")))
            log_to_ui(result_text_area, "'근무시간현황' 메뉴 클릭.")
            commute_menu.click()

            # 날짜 입력창이 나타날 때까지 대기 (CSS 선택자 기준)
            #log_to_ui(result_text_area, "날짜 입력창을 기다립니다...")
            date_pickers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".OBTDatePickerRebuild_inputYMD__PtxMy.OBTDatePickerRebuild_dateInput__35pTn")))
            #log_to_ui(result_text_area, "날짜 입력창 확인 성공")
            if len(date_pickers) < 2:
                raise Exception("날짜 입력창 2개를 찾는 데 실패했습니다.")

            start_date_picker = date_pickers[0]
            end_date_picker = date_pickers[1]

            # 키보드 조합(Ctrl+A, Backspace)으로 날짜를 강제 삭제하고 입력
            log_to_ui(result_text_area, f"시작일자 {start_date} 입력 중...")
            start_date_picker.click()
            time.sleep(0.3)
            start_date_picker.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            start_date_picker.send_keys(Keys.BACK_SPACE)
            time.sleep(0.3)
            start_date_picker.send_keys(start_date)
            time.sleep(0.3)
            log_to_ui(result_text_area, f"종료일자 {end_date} 입력 중...")
            end_date_picker.click()
            time.sleep(0.3)
            end_date_picker.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            end_date_picker.send_keys(Keys.BACK_SPACE)
            time.sleep(0.3)
            end_date_picker.send_keys(end_date)
            
            log_to_ui(result_text_area, "기존 네트워크 기록을 삭제합니다...")
            del driver.requests
            time.sleep(1)
            # Enter 키 입력으로 날짜 갱신
            log_to_ui(result_text_area, "Enter 키를 눌러 날짜를 적용(자동 갱신)합니다...")
            end_date_picker.send_keys(Keys.ENTER)
            # -----------------------------------------------------------------

        except Exception as e:
            log_to_ui(result_text_area, f"오류 발생! 날짜 설정 또는 자동 갱신에 실패했습니다.")
            driver.quit()
            messagebox.showerror("페이지 이동/조회 실패", f"날짜 입력 또는 자동 갱신(Enter)에 실패했습니다.\n오류: {e}")
            return
            # -----------------------------------------------------------------
        
        
        # 5. 'selectTab2' 네트워크 요청 가로채기
        try:
            log_to_ui(result_text_area, f"'{INTERCEPT_URL_KEYWORD}' 요청 가로채기를 시작합니다 (최대 30초 대기)...")
            request = driver.wait_for_request(INTERCEPT_URL_KEYWORD, timeout=30)
            log_to_ui(result_text_area, "... 'selectTab2' 요청 가로채기 성공!")
            response = request.response
            if not response:
                raise Exception("서버에서 응답을 받지 못했습니다.")
            response_body_bytes = response.body
            encoding = response.headers.get('Content-Encoding', '')
            if 'gzip' in encoding:
                log_to_ui(result_text_area, "Gzip 압축 감지. 압축을 해제합니다...")
                decompressed_body = gzip.decompress(response_body_bytes)
                response_text = decompressed_body.decode('utf-8')
            else:
                log_to_ui(result_text_area, "압축 없음. UTF-8로 디코딩합니다...")
                response_text = response_body_bytes.decode('utf-8')
            if not response_text:
                raise Exception("데이터를 디코딩했으나, 텍스트가 비어있습니다.")
            json_data = json.loads(response_text)
            data_list = json_data[JSON_DATA_LIST_KEY]

        except Exception as e:
            driver.quit()
            messagebox.showerror("데이터 수신 실패", f"'{INTERCEPT_URL_KEYWORD}' 요청을 가로CSS으나 데이터를 처리하지 못했습니다.\n('조회' 버튼 클릭 시 이 요청이 발생하는지 확인하세요)\n오류: {e}")
            return
        
        driver.quit() 
        log_to_ui(result_text_area, f"\n데이터 수신 성공! {len(data_list)}개의 기록을 바탕으로 계산을 시작합니다...")

        # 8. 계산 함수 호출
        result = calculate_work_hours(data_list, result_text_area)

        # 9. 최종 결과 표시
        log_to_ui(result_text_area, "\n" + result)

    except Exception as e:
        try:
            driver.quit()
        except:
            pass
        messagebox.showerror("Error", f"자동화 중 알 수 없는 오류 발생: {e}")

# --- (4) UI 설정 ---

# 디자인 색상 및 폰트 정의
BG_COLOR = "#2E2E2E"
LOG_BG_COLOR = "#1E1E1E"
TEXT_COLOR = "#EAEAEA"
BUTTON_COLOR = "#007ACC"
BUTTON_TEXT_COLOR = "#FFFFFF"
BUTTON_ACTIVE_COLOR = "#005C99"
ENTRY_BG_COLOR = "#3C3C3C"
ENTRY_TEXT_COLOR = "#EAEAEA"

APP_FONT = ("Malgun Gothic", 12, "bold")
BUTTON_FONT = ("Malgun Gothic", 12, "bold")

# ID/PW 저장을 위한 파일 이름 정의
def get_script_directory():
    if getattr(sys, 'frozen', False):
        # .exe로 실행된 경우 (pyinstaller)
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        # .py 스크립트로 실행된 경우
        return os.path.dirname(os.path.abspath(__file__))

# ID/PW 저장을 위한 절대 경로 정의
script_dir = get_script_directory()
LOGIN_FILE = os.path.join(script_dir, "login_info.txt")

# ID/PW 저장 함수
def save_credentials(user_id, user_pw):
    try:
        with open(LOGIN_FILE, "w", encoding="utf-8") as f:
            f.write(f"{user_id}\n")
            f.write(f"{user_pw}\n")
    except Exception as e:
        # 파일 저장 실패가 크리티컬한 문제는 아니므로, 콘솔에만 오류를 출력합니다.
        print(f"로그인 정보 저장 실패: {e}")

# ID/PW 불러오기 함수
def load_credentials():
    print("--- 로그인 정보 불러오기 시도 ---") # 디버깅용
    print(f"찾는 파일 경로: {LOGIN_FILE}") # 디버깅용
    
    if not os.path.exists(LOGIN_FILE):
        print("결과: 'login_info.txt' 파일을 찾을 수 없습니다.") # 디버깅용
        return "", "" # 파일이 없으면 빈 값 반환
    try:
        with open(LOGIN_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
            if len(lines) >= 2:
                user_id = lines[0].strip()
                user_pw = lines[1].strip()
                print(f"결과: 성공. ID={user_id}") # 디버깅용
                return user_id, user_pw
            else:
                print(f"결과: 파일은 있으나, 내용(라인 수)이 부족합니다. ({len(lines)}줄)") # 디버깅용
                return "", "" # 파일 내용이 비정상이면 빈 값 반환
    except Exception as e:
        print(f"결과: 파일 읽기 중 오류 발생: {e}")
        return "", ""
    
# 버튼 클릭 시 ID/PW 저장 로직 
def on_button_click():
    # 1. UI에서 ID/PW/날짜 가져오기
    user_id = id_entry.get()
    user_pw = pw_entry.get()
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    # 2. 입력값 비어있는지 확인
    if not all([user_id, user_pw, start_date, end_date]):
        messagebox.showwarning("입력 오류", "아이디, 비밀번호, 시작일, 종료일을 모두 입력해야 합니다.")
        return
    
    # 입력된 ID/PW를 파일에 저장
    save_credentials(user_id, user_pw)
    
    # 날짜 형식(YYYY-MM-DD) 검증
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        messagebox.showwarning("입력 오류", "날짜 형식이 올바르지 않습니다.\n(YYYY-MM-DD 형식으로 입력하세요)")
        return

    # 3. 스레드로 자동화 함수 실행 (UI 멈춤 방지)
    import threading
    threading.Thread(target=run_automation_and_calculate, 
                     args=(user_id, user_pw, start_date, end_date, result_text_area), 
                     daemon=True).start()

window = tk.Tk()
window.title("초과근무 시간 계산기 (v0.4.3)")
window.geometry("600x700")
window.attributes('-topmost', True)
window.config(bg=BG_COLOR)

# UI 생성 전, 저장된 ID/PW 불러오기
loaded_id, loaded_pw = load_credentials()
id_var = tk.StringVar(window)
pw_var = tk.StringVar(window)
id_var.set(loaded_id)
pw_var.set(loaded_pw)

# --- ID/PW 입력을 위한 프레임 및 위젯 ---
login_frame = tk.Frame(window, bg=BG_COLOR)
login_frame.pack(pady=(20, 0), padx=20, fill=tk.X)
id_label = tk.Label(login_frame, text="아이디:", font=APP_FONT, bg=BG_COLOR, fg=TEXT_COLOR, width=8, anchor='w')
id_label.pack(side=tk.LEFT, padx=(0, 10))
id_entry = tk.Entry(login_frame, font=APP_FONT, bg=ENTRY_BG_COLOR, fg=ENTRY_TEXT_COLOR, insertbackground=TEXT_COLOR, relief=tk.FLAT, borderwidth=0, textvariable=id_var)
id_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

pw_frame = tk.Frame(window, bg=BG_COLOR)
pw_frame.pack(pady=10, padx=20, fill=tk.X)
pw_label = tk.Label(pw_frame, text="비밀번호:", font=APP_FONT, bg=BG_COLOR, fg=TEXT_COLOR, width=8, anchor='w')
pw_label.pack(side=tk.LEFT, padx=(0, 10))
pw_entry = tk.Entry(pw_frame, font=APP_FONT, show="*", bg=ENTRY_BG_COLOR, fg=ENTRY_TEXT_COLOR, insertbackground=TEXT_COLOR, relief=tk.FLAT, borderwidth=0, textvariable=pw_var)
pw_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

# --- 날짜 입력을 위한 프레임 및 위젯 ---
today_str = datetime.now().strftime("%Y-%m-%d")
first_day_of_month_str = datetime.now().replace(day=1).strftime("%Y-%m-%d")

date_frame_start = tk.Frame(window, bg=BG_COLOR)
date_frame_start.pack(pady=(5, 0), padx=20, fill=tk.X)
start_label = tk.Label(date_frame_start, text="시작일:", font=APP_FONT, bg=BG_COLOR, fg=TEXT_COLOR, width=8, anchor='w')
start_label.pack(side=tk.LEFT, padx=(0, 10))
start_date_entry = tk.Entry(date_frame_start, font=APP_FONT, bg=ENTRY_BG_COLOR, fg=ENTRY_TEXT_COLOR, insertbackground=TEXT_COLOR, relief=tk.FLAT, borderwidth=0)
start_date_entry.insert(0, first_day_of_month_str)
start_date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

date_frame_end = tk.Frame(window, bg=BG_COLOR)
date_frame_end.pack(pady=10, padx=20, fill=tk.X)
end_label = tk.Label(date_frame_end, text="종료일:", font=APP_FONT, bg=BG_COLOR, fg=TEXT_COLOR, width=8, anchor='w')
end_label.pack(side=tk.LEFT, padx=(0, 10))
end_date_entry = tk.Entry(date_frame_end, font=APP_FONT, bg=ENTRY_BG_COLOR, fg=ENTRY_TEXT_COLOR, insertbackground=TEXT_COLOR, relief=tk.FLAT, borderwidth=0)
end_date_entry.insert(0, today_str)
end_date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

# 시작 버튼
process_button = tk.Button(window, text="자동 계산 시작하기", font=BUTTON_FONT, command=on_button_click, bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, activebackground=BUTTON_ACTIVE_COLOR, activeforeground=BUTTON_TEXT_COLOR, relief=tk.FLAT, borderwidth=0, padx=20)
process_button.pack(pady=10, padx=20, fill=tk.X, ipady=8)

# 로그 표시 영역
result_text_area = scrolledtext.ScrolledText(window, wrap=tk.WORD, font=APP_FONT, bg=LOG_BG_COLOR, fg=TEXT_COLOR, relief=tk.FLAT, borderwidth=0, insertbackground=TEXT_COLOR, state=tk.DISABLED)
result_text_area.pack(pady=(0, 20), padx=20, fill=tk.BOTH, expand=True)

window.mainloop()