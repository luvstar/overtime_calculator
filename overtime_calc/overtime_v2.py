import tkinter as tk
from tkinter import scrolledtext, messagebox
from tkinter import ttk
import pandas as pd
import time
import holidays
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
START_TIME_KEY = "comeTm" # 실출근
END_TIME_KEY = "leaveTm" # 실퇴근

# --- UI 로깅 함수 ---
def log_to_ui(text_widget, message):
    log_bg_color = "#1E1E1E"
    text_color = "#EAEAEA"
    text_widget.config(state=tk.NORMAL, bg=log_bg_color, fg=text_color)
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END) 
    text_widget.config(state=tk.DISABLED, bg=log_bg_color, fg=text_color)
    window.update_idletasks() 

# --- UI 상태 업데이트 함수 (프로그레스바 & 라벨) ---
def update_status(progress_val, message_text):
    try:
        progress_bar['value'] = progress_val
        status_label.config(text=message_text)
        window.update_idletasks()
    except Exception:
        pass 

# --- 결과 텍스트 출력 함수 ---
def show_result(text_widget, message):
    text_widget.config(state=tk.NORMAL)
    text_widget.delete('1.0', tk.END) 
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END) 
    text_widget.config(state=tk.DISABLED)

# --- (2) 근무 시간 계산 로직 (출장 시 편측 기록 보정: 09시/18시) ---
def calculate_work_hours(json_data_list, text_widget):
    def format_timedelta_simple(td):
        if pd.isna(td): return "00:00:00"
        total_seconds = int(abs(td.total_seconds()))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    APP_START_KEY = "appcomeTm"
    APP_END_KEY = "appEndTm"
    STATUS_KEY = "atNm"
    
    # UI 초기화 및 태그 설정
    text_widget.config(state=tk.NORMAL)
    text_widget.delete('1.0', tk.END)
    
    text_widget.tag_config("warning", foreground="#FF6B6B")
    text_widget.tag_config("success", foreground="#51CF66")
    text_widget.tag_config("default", foreground="#EAEAEA")
    
    try:
        if not json_data_list:
            text_widget.insert(tk.END, "오류: JSON 데이터 목록이 비어있습니다.\n")
            text_widget.config(state=tk.DISABLED)
            return
        
        df = pd.DataFrame(json_data_list)
        df['날짜_dt'] = pd.to_datetime(df[DATE_KEY])
        df['날짜'] = df['날짜_dt'].dt.date

        unique_years = df['날짜_dt'].dt.year.unique().tolist()
        kr_holidays = holidays.KR(years=unique_years)

        def format_hhmm(hhmm_val):
            if pd.isna(hhmm_val) or hhmm_val == '': return None
            try: hhmm_str = str(int(float(hhmm_val))).zfill(4) 
            except ValueError: hhmm_str = str(hhmm_val)
            if len(hhmm_str) < 4: return None
            return f"{hhmm_str[:2]}:{hhmm_str[2:]}"

        df['인정출근'] = df[APP_START_KEY].apply(format_hhmm)
        df['인정퇴근'] = df[APP_END_KEY].apply(format_hhmm)
        
        # --- 핵심 로직 ---
        def get_actual_work_time(row):
            curr_date = row['날짜']
            is_weekend = curr_date.weekday() >= 5 
            status_val = str(row.get(STATUS_KEY, ""))
            
            # [1순위] 특수 상태 (연차, 휴가, 출장 등)
            special_keywords = ["연차", "휴가", "경조", "출장", "반차", "공가"]
            if any(keyword in status_val for keyword in special_keywords):
                
                # --- [수정됨] 출장 시 실 근무시간 보정 로직 ---
                if "출장" in status_val:
                    has_start = pd.notna(row[START_TIME_KEY])
                    has_end = pd.notna(row[END_TIME_KEY])
                    
                    act_start = None
                    act_end = None
                    
                    # Case 1: 출/퇴근 모두 있음
                    if has_start and has_end:
                        act_start = pd.to_datetime(f"{curr_date} {row[START_TIME_KEY]}")
                        act_end = pd.to_datetime(f"{curr_date} {row[END_TIME_KEY]}")
                    
                    # Case 2: 출근만 있음 -> 퇴근 18:00 가정
                    elif has_start and not has_end:
                        act_start = pd.to_datetime(f"{curr_date} {row[START_TIME_KEY]}")
                        act_end = pd.to_datetime(f"{curr_date} 18:00:00")
                        
                    # Case 3: 퇴근만 있음 -> 출근 09:00 가정
                    elif not has_start and has_end:
                        act_start = pd.to_datetime(f"{curr_date} 09:00:00")
                        act_end = pd.to_datetime(f"{curr_date} {row[END_TIME_KEY]}")
                    
                    # 계산 가능한 시간이 만들어졌다면 확인
                    if act_start and act_end:
                        act_duration = (act_end - act_start) - pd.Timedelta(hours=1) # 휴게 차감
                        # 보정된 실 근무시간이 8시간을 넘으면, 이걸 인정
                        if act_duration > pd.Timedelta(hours=8):
                            return act_duration

                # 위 조건(출장 8시간 초과)에 해당하지 않으면 기존 로직(인정 시간) 사용
                if pd.notna(row['인정출근']) and pd.notna(row['인정퇴근']):
                    start_dt = pd.to_datetime(f"{curr_date} {row['인정출근']}")
                    end_dt = pd.to_datetime(f"{curr_date} {row['인정퇴근']}")
                    work_time = (end_dt - start_dt) - pd.Timedelta(hours=1)
                    return max(work_time, pd.Timedelta(0))
                else:
                    if "반차" in status_val: return pd.Timedelta(hours=4)
                    return pd.Timedelta(hours=8)

            # [2순위] 실 근무 기록 존재 (일반 근무)
            if pd.notna(row[START_TIME_KEY]) and pd.notna(row[END_TIME_KEY]):
                start_dt = pd.to_datetime(f"{row['날짜']} {row[START_TIME_KEY]}")
                end_dt = pd.to_datetime(f"{row['날짜']} {row[END_TIME_KEY]}")
                if end_dt < start_dt:
                    end_dt += pd.Timedelta(days=1)
                work_time = (end_dt - start_dt) - pd.Timedelta(hours=1)
                return max(work_time, pd.Timedelta(0))

            # [3순위] 주말 기록 없음 -> 제외
            if is_weekend: 
                return pd.NaT

            # [4순위] 평일 공휴일 -> 0시간 (목표 차감용)
            if curr_date in kr_holidays: 
                return pd.Timedelta(0)
            
            # [5순위] 그 외 -> 제외
            return pd.NaT

        df['실근무시간'] = df.apply(get_actual_work_time, axis=1)
        
        df = df.dropna(subset=['실근무시간'])
        
        if df.empty:
            text_widget.insert(tk.END, "표시할 근무 데이터가 없습니다.\n", "default")
            text_widget.config(state=tk.DISABLED)
            return
        
        df['주차'] = df['날짜_dt'].dt.isocalendar().week
        weekly_groups = df.groupby('주차')
        
        text_widget.insert(tk.END, "=== 📅 일별 근무 현황 (미근무 주말 제외) ===\n", "default")
        
        for _, row in df.iterrows():
            status_val = str(row.get(STATUS_KEY, ""))
            tag = ""
            holiday_name = kr_holidays.get(row['날짜'])
            #if holiday_name: tag = f" [공휴일:{holiday_name}]"
            if any(k in status_val for k in ["연차", "반차", "출장", "휴가"]): tag = f" [{status_val}]"
            
            line = f"[{row['날짜']}{tag}] 실근무: {format_timedelta_simple(row['실근무시간'])}\n"
            text_widget.insert(tk.END, line, "default")
            
        text_widget.insert(tk.END, "\n=== 📊 주별 요약 (유동적 목표 시간) ===\n", "default")
        
        for week_num, group in weekly_groups:
            total_work = group['실근무시간'].sum()
            
            start_date_str = group['날짜_dt'].min().strftime("%m.%d")
            end_date_str = group['날짜_dt'].max().strftime("%m.%d")
            date_range_str = f"({start_date_str}~{end_date_str})"

            holiday_count = 0
            for date_val in group['날짜']:
                if date_val in kr_holidays and date_val.weekday() < 5:
                    holiday_count += 1
            
            target_hours = pd.Timedelta(hours=40) - pd.Timedelta(hours=8 * holiday_count)
            target_str = f"{int(target_hours.total_seconds()//3600)}H"
            
            if total_work >= target_hours:
                diff = total_work - target_hours
                status_str = f"✅ 초과: {format_timedelta_simple(diff)}"
                tag_name = "success"
            else:
                diff = target_hours - total_work
                status_str = f"⚠️ 미달: {format_timedelta_simple(diff)}"
                tag_name = "warning"
                
            line = f"[{week_num}주차 {date_range_str} | 목표 {target_str}] 총 근무: {format_timedelta_simple(total_work)} | {status_str}\n"
            text_widget.insert(tk.END, line, tag_name)
            
        text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)

    except Exception as e:
        import traceback
        traceback.print_exc()
        text_widget.insert(tk.END, f"계산 중 오류 발생: {e}\n", "warning")
        text_widget.config(state=tk.DISABLED)

# --- (3) Selenium 자동화 로직 ---
def run_automation_and_calculate(user_id, user_pw, start_date, end_date, result_text_area):
    try:
        # 0. UI 업데이트
        result_text_area.config(state=tk.NORMAL)
        result_text_area.delete('1.0', tk.END) 
        result_text_area.config(state=tk.DISABLED)
        show_result(result_text_area, "") # 결과창 비우기
        update_status(5, "브라우저 실행 중...")

        # 헤드리스 모드(백그라운드 실행) 옵션 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options) 

        update_status(15, "로그인 페이지 접속 중...")
        driver.get(LOGIN_URL)
        time.sleep(3) 
        
        # 3. 자동 로그인
        try:
            wait = WebDriverWait(driver, 10)
            update_status(15, "로그인 페이지 접속 중...")
            id_input = wait.until(EC.presence_of_element_located((By.ID, "reqLoginId")))
            update_status(25, "아이디 입력 중...")
            id_input.send_keys(user_id)
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '다음')]")))
            next_button.click()
            update_status(35, "비밀번호 입력 중...")
            pw_input = wait.until(EC.element_to_be_clickable((By.ID, "reqLoginPw")))
            pw_input.send_keys(user_pw)
            login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '로그인')]")))
            update_status(45, "로그인 시도 중...")
            login_button.click()
        except Exception as e:
            update_status(0, "로그인 실패")
            driver.quit()
            messagebox.showerror("로그인 실패", f"로그인 요소를 찾을 수 없습니다.\n오류: {e}")
            return

        update_status(55, "로그인 성공. 데이터 페이지로 이동 중...")

        # --- 4. 데이터 페이지 이동 및 날짜 설정 ---
        try:
            wait = WebDriverWait(driver, 10) 
            commute_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-name='근무시간현황']")))
            update_status(60, "근무 현황 페이지 이동 중...")
            commute_menu.click()

            date_pickers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".OBTDatePickerRebuild_inputYMD__PtxMy.OBTDatePickerRebuild_dateInput__35pTn")))
            if len(date_pickers) < 2:
                raise Exception("날짜 입력창 2개를 찾는 데 실패했습니다.")

            start_date_picker = date_pickers[0]
            end_date_picker = date_pickers[1]

            update_status(65, f"날짜 입력 중 ({start_date} ~ {end_date})...")
            start_date_picker.click()
            time.sleep(0.3)
            start_date_picker.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            start_date_picker.send_keys(Keys.BACK_SPACE)
            time.sleep(0.3)
            start_date_picker.send_keys(start_date)
            time.sleep(0.3)
            update_status(75, f"날짜 입력 중 ({start_date} ~ {end_date})...")
            end_date_picker.click()
            time.sleep(0.3)
            end_date_picker.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            end_date_picker.send_keys(Keys.BACK_SPACE)
            time.sleep(0.3)
            end_date_picker.send_keys(end_date)
            
            update_status(77, "기존 네트워크 기록을 삭제 중...")
            del driver.requests
            time.sleep(1)
            update_status(80, "네트워크 기록 갱신 중...")
            end_date_picker.send_keys(Keys.ENTER)

        except Exception as e:
            update_status(0, "오류 발생! 날짜 설정 또는 자동 갱신에 실패했습니다.")
            driver.quit()
            messagebox.showerror("페이지 이동/조회 실패", f"날짜 입력 또는 자동 갱신(Enter)에 실패했습니다.\n오류: {e}")
            return
        
        # 5. 'selectTab2' 네트워크 요청 가로채기
        try:
            update_status(82, "네트워크 패킷 데이터 수신 중...")
            request = driver.wait_for_request(INTERCEPT_URL_KEYWORD, timeout=30)
            update_status(85, "네트워크 패킷 데이터 수신 성공")
            response = request.response
            if not response:
                raise Exception("서버에서 응답을 받지 못했습니다.")
            response_body_bytes = response.body
            encoding = response.headers.get('Content-Encoding', '')
            if 'gzip' in encoding:
                update_status(88, "네트워크 패킷 데이터 압축 해제 진행 중...")
                decompressed_body = gzip.decompress(response_body_bytes)
                response_text = decompressed_body.decode('utf-8')
            else:
                update_status(88, "네트워크 패킷 데이터 UTF-8 디코딩 진행 중...")
                response_text = response_body_bytes.decode('utf-8')
            if not response_text:
                raise Exception("데이터를 디코딩했으나, 텍스트가 비어있습니다.")
            json_data = json.loads(response_text)
            data_list = json_data[JSON_DATA_LIST_KEY]

        except Exception as e:
            driver.quit()
            messagebox.showerror("데이터 수신 실패", f"'{INTERCEPT_URL_KEYWORD}' 요청을 가로챘으나 데이터를 처리하지 못했습니다.\n오류: {e}")
            return
        
        driver.quit() 
        update_status(95, f"\n데이터 수신 성공! {len(data_list)}개의 기록 기반 근무시간 계산 중...")

        # 8. 계산 함수 호출
        #result = calculate_work_hours(data_list, result_text_area)

        # 9. 최종 결과 표시
        update_status(98, f"계산완료. 결과 출력 중...")
        time.sleep(0.5)
        #log_to_ui(result_text_area, "\n" + result)
        calculate_work_hours(data_list, result_text_area)
        update_status(100, f"작업 완료")

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
STATUS_TEXT_COLOR = "#AAAAAA" 

APP_FONT = ("Malgun Gothic", 12, "bold")
BUTTON_FONT = ("Malgun Gothic", 12, "bold")
STATUS_FONT = ("Malgun Gothic", 10)

def get_script_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

script_dir = get_script_directory()
LOGIN_FILE = os.path.join(script_dir, "login_info.txt")

def save_credentials(user_id, user_pw):
    try:
        with open(LOGIN_FILE, "w", encoding="utf-8") as f:
            f.write(f"{user_id}\n")
            f.write(f"{user_pw}\n")
    except Exception as e:
        print(f"로그인 정보 저장 실패: {e}")

def load_credentials():
    if not os.path.exists(LOGIN_FILE):
        return "", ""
    try:
        with open(LOGIN_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
            if len(lines) >= 2:
                return lines[0].strip(), lines[1].strip()
            return "", "" 
    except Exception as e:
        return "", ""
    
def on_button_click(event=None):
    user_id = id_entry.get()
    user_pw = pw_entry.get()
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    if not all([user_id, user_pw, start_date, end_date]):
        messagebox.showwarning("입력 오류", "아이디, 비밀번호, 시작일, 종료일을 모두 입력해야 합니다.")
        return
    
    save_credentials(user_id, user_pw)
    
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        messagebox.showwarning("입력 오류", "날짜 형식이 올바르지 않습니다.\n(YYYY-MM-DD 형식으로 입력하세요)")
        return

    import threading
    threading.Thread(target=run_automation_and_calculate, 
                     args=(user_id, user_pw, start_date, end_date, result_text_area), 
                     daemon=True).start()

window = tk.Tk()
window.title("초과근무 시간 계산기 (v2.0)")
window.geometry("600x720")
window.attributes('-topmost', True)
window.config(bg=BG_COLOR)

window.bind('<Return>', on_button_click)

loaded_id, loaded_pw = load_credentials()
id_var = tk.StringVar(window)
pw_var = tk.StringVar(window)
id_var.set(loaded_id)
pw_var.set(loaded_pw)

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

process_button = tk.Button(window, text="자동 계산 시작하기(ENTER)", font=BUTTON_FONT, command=on_button_click, bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, activebackground=BUTTON_ACTIVE_COLOR, activeforeground=BUTTON_TEXT_COLOR, relief=tk.FLAT, borderwidth=0, padx=20)
process_button.pack(pady=10, padx=20, fill=tk.X, ipady=8)

status_label = tk.Label(window, text="대기 중...", font=STATUS_FONT, bg=BG_COLOR, fg=STATUS_TEXT_COLOR)
status_label.pack(pady=(10, 5), padx=20, anchor='w')
style = ttk.Style()
style.theme_use('default')
style.configure("TProgressbar", thickness=15, troughcolor=ENTRY_BG_COLOR, background=BUTTON_COLOR)
progress_bar = ttk.Progressbar(window, style="TProgressbar", orient="horizontal", length=100, mode="determinate")
progress_bar.pack(pady=(0, 15), padx=20, fill=tk.X)

result_text_area = scrolledtext.ScrolledText(window, wrap=tk.WORD, font=APP_FONT, bg=LOG_BG_COLOR, fg=TEXT_COLOR, relief=tk.FLAT, borderwidth=0, insertbackground=TEXT_COLOR, state=tk.DISABLED)
result_text_area.pack(pady=(0, 20), padx=20, fill=tk.BOTH, expand=True)

window.mainloop()



