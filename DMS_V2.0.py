import sys
import time
import os
import sys
import tkinter as tk
import shutil
import tempfile
from tkinter import scrolledtext, messagebox, StringVar, BooleanVar, Checkbutton
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoAlertPresentException, StaleElementReferenceException, UnexpectedAlertPresentException, JavascriptException
from concurrent.futures import ThreadPoolExecutor
from PIL import Image, ImageTk

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

def resource_path(relative_path):
    """獲取打包後的資源檔案路徑 (確保 PyInstaller 打包後可以找到資源)"""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path

def get_icon_path():
    """返回圖示的路徑"""
    icon_filename = "002.ico"
    if hasattr(sys, "_MEIPASS"):
        # 打包後從臨時目錄中獲取
        return os.path.join(sys._MEIPASS, icon_filename)
    else:
        # 開發模式下的完整路徑
        return r"C:\Users\a0934\AppData\Local\Programs\Python\Python312\002.ico"

def choose_browser():
    """顯示選擇瀏覽器的 GUI 按鈕"""
    root = tk.Tk()
    root.title("DMS 自動簽核工具_V2.0")
    root.geometry("350x250")  # 視窗大小
    root.resizable(False, False)  # 禁止調整視窗大小
    root.iconbitmap(get_icon_path())

    # **載入圖示**
    chrome_icon_path = resource_path("chrome.ico")
    edge_icon_path = resource_path("edge.ico")

    try:
        # ✅ 先用 PIL 開啟 .ico，然後轉成 PhotoImage
        chrome_img = Image.open(chrome_icon_path).convert("RGBA")
        edge_img = Image.open(edge_icon_path).convert("RGBA")

        chrome_img = chrome_img.resize((48, 48))
        edge_img = edge_img.resize((48, 48))

        chrome_icon = ImageTk.PhotoImage(chrome_img)
        edge_icon = ImageTk.PhotoImage(edge_img)

    except FileNotFoundError:
        print(f"❌ 找不到圖示: {chrome_icon_path} 或 {edge_icon_path}")
        sys.exit(1)

    selected_browser = tk.StringVar()

    def select_chrome():
        selected_browser.set("chrome")
        root.destroy()

    def select_edge():
        selected_browser.set("edge")
        root.destroy()

    # **標題**
    label = tk.Label(root, text="請選擇要使用的瀏覽器", font=("Arial", 16, "bold"))
    label.pack(pady=10)

    # **按鈕容器**
    frame = tk.Frame(root)
    frame.pack(pady=10)

    # **Google Chrome 按鈕**
    btn_chrome = tk.Button(frame, text=" Google Chrome", font=("Arial", 14), fg="black", bg="white",
                            width=200, height=40, image=chrome_icon, compound="left", anchor="w", padx=20, command=select_chrome)
    btn_chrome.image = chrome_icon  # 防止垃圾回收
    btn_chrome.grid(row=0, column=0, padx=20, pady=10)  # 增加按鈕間距

    # **Microsoft Edge 按鈕**
    btn_edge = tk.Button(frame, text=" Microsoft Edge", font=("Arial", 14), fg="black", bg="white", width=200, height=40, image=edge_icon
                        , compound="left", anchor="w", padx=20, command=select_edge)
    btn_edge.image = edge_icon
    btn_edge.grid(row=1, column=0, padx=20, pady=10)  # 增加按鈕間距

    root.mainloop()
    return selected_browser.get()

def get_driver_path(driver_name):
    """獲取 WebDriver 路徑，確保 Selenium 可以讀取"""
    if getattr(sys, 'frozen', False):  # 如果是 PyInstaller 打包後的 .exe
        temp_dir = tempfile.gettempdir()
        extracted_driver_path = os.path.join(temp_dir, driver_name)

        bundled_driver_path = os.path.join(sys._MEIPASS, driver_name)
        if not os.path.exists(bundled_driver_path):
            print(f"❌ WebDriver {driver_name} 並未打包到 .exe 內！")
            sys.exit(1)

        # 確保 `temp_dir` 內的 WebDriver 可執行
        try:
            shutil.copy(bundled_driver_path, extracted_driver_path)
            os.chmod(extracted_driver_path, 0o755)  # 確保執行權限
        except Exception as e:
            print(f"⚠️ WebDriver 複製失敗: {e}")
            sys.exit(1)

        return extracted_driver_path

    else:
        # 直接使用系統上的 WebDriver
        return f"C:/{'chromedriver' if driver_name == 'chromedriver.exe' else 'edgedriver'}/{driver_name}"


# **動態選擇瀏覽器**
selected_browser = choose_browser()
if not selected_browser:
    sys.exit("🚫 使用者取消操作")  # 如果未選擇，則退出程式

# ✅ 設定 WebDriver 路徑
CHROME_DRIVER_PATH = get_driver_path("chromedriver.exe")
EDGE_DRIVER_PATH = get_driver_path("msedgedriver.exe")

# ✅ 顯示確認
print(f"✅ ChromeDriver 路徑: {CHROME_DRIVER_PATH}")
print(f"✅ EdgeDriver 路徑: {EDGE_DRIVER_PATH}")

# **瀏覽器選項設定**
def set_browser_options(browser_type):
    """為 Chrome 和 Edge 設定最佳化選項"""
    options = ChromeOptions() if browser_type == "chrome" else EdgeOptions()
    options.add_argument("--disable-features=PaintHolding")  # 關閉 Paint 渲染
    options.add_argument("--disable-logging")  # 關閉不必要的日誌
    options.add_argument("--log-level=3")  # 降低日誌輸出等級，避免干擾
    
    # ✅ Edge 需要這個選項
    if browser_type == "edge":
        options.add_argument("--remote-debugging-port=0")
    return options

# 啟動瀏覽器
if selected_browser == "chrome":
    options = set_browser_options("chrome")
    service = ChromeService(get_driver_path("chromedriver.exe"))
    driver = webdriver.Chrome(service=service, options=options)
    print("✅ 使用 Google Chrome")
else:
    options = set_browser_options("edge")
    service = EdgeService(get_driver_path("msedgedriver.exe"))
    driver = webdriver.Edge(service=service, options=options)  # 修正
    print("✅ 使用 Microsoft Edge")



# **開啟 DMS 網站**
driver.get("https://dmsweb.inventec.com/be/my_work_flow/")
# **建立 GUI 介面**
root = tk.Tk()
root.title("DMS 自動簽核工具_V2.0")
root.geometry("495x415")
root.iconbitmap(get_icon_path())

# **輸出文字框**
output_text = scrolledtext.ScrolledText(root, width=65, height=15)
output_text.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

def log_message(message):
    """輸出內容至 GUI 文字框，只保留最新 20 行"""
    output_text.insert(tk.END, message + "\n")
    output_text.see(tk.END)

    # **只保留最新 20 行**
    lines = output_text.get("1.0", tk.END).split("\n")
    if len(lines) > 20:
        output_text.delete("1.0", f"{len(lines) - 20}.0")

    root.update_idletasks()

log_message("✅ 已開啟 Chrome，請手動登入DMS，並確保沒有其他分頁，完成後按下『開始簽核』按鈕")

# **簽核狀態標題**
status_label = tk.Label(root, text="簽呈狀態選擇：", font=("Arial", 12, "bold"))
status_label.grid(row=1, column=0, columnspan=3, padx=10, pady=(5, 2), sticky="w")

# **選擇簽核類別**
status_options = ["簽核中", "進行中", "修模中", "已修模"]
status_vars = {status: BooleanVar() for status in status_options}

# **優化選擇框間距，讓它們對齊並居中**
status_frame = tk.Frame(root)
status_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=2)

for i, status in enumerate(status_options):
    check = Checkbutton(status_frame, text=status, variable=status_vars[status])
    if status in ["簽核中", "進行中"]:
        check.config(fg="red")
    check.pack(side="left", padx=10)

# **簽核內容標題**
input_label = tk.Label(root, text="簽核內容：", font=("Arial", 12, "bold"))
input_label.grid(row=3, column=0, padx=10, pady=(10, 2), sticky="w")

# **輸入簽核內容**
input_text_var = StringVar()
input_text_var.set("EC OK")
input_entry = tk.Entry(root, textvariable=input_text_var, width=40)
input_entry.grid(row=4, column=0, columnspan=3, padx=10, pady=2)

# **開始簽核按鈕**
def start_processing():
    selected_statuses = [status for status, var in status_vars.items() if var.get()]
    if not selected_statuses:
        messagebox.showwarning("錯誤", "請選擇至少一個簽核狀態")
        return

    # **"簽核中" 和 "進行中" 需要額外確認**
    if "簽核中" in selected_statuses or "進行中" in selected_statuses:
        confirm = messagebox.askyesno("確認", "⚠️ 你選擇了「簽核中」或「進行中」，確定要執行嗎？")
        if not confirm:
            return

    log_message(f"✅ 開始處理簽核類別: {', '.join(selected_statuses)}")
    process_signatures(selected_statuses, input_text_var.get())
    
    # **使用多線程執行簽核，提高效率**
    executor = ThreadPoolExecutor(max_workers=2)
    executor.submit(process_signatures, selected_statuses, input_text_var.get())

start_button = tk.Button(root, text="開始簽核", command=start_processing, font=("Arial", 12, "bold"), fg="blue", width=12, height=2)
start_button.grid(row=5, column=0, padx=10, pady=10, sticky="w")

# **關閉按鈕**
def close_program():
    driver.quit()
    root.quit()

close_button = tk.Button(root, text="關閉", command=close_program, font=("Arial", 12, "bold"), width=12, height=2)
close_button.grid(row=5, column=2, padx=10, pady=10, sticky="e")

def process_signatures(selected_statuses, input_text):
    """處理簽核"""
    driver.get("https://dmsweb.inventec.com/be/my_work_flow/")
    wait = WebDriverWait(driver, 10)

    def get_signatures():
        time.sleep(2)
        xpath_query = " | ".join([f"//td//span[contains(text(), '{status}')]" for status in selected_statuses])
        return driver.find_elements(By.XPATH, xpath_query)

    while True:
        elements = get_signatures()
        if not elements:
            log_message("🎉 所有選擇的簽呈已完成簽核")
            break

        log_message(f"🔍 找到 {len(elements)} 份待處理簽呈")

        for element in elements:
            try:
                # **使用 JavaScript 點擊**
                driver.execute_script("arguments[0].click();", element)
                log_message("✅ 點擊成功，進入簽核頁面")
                time.sleep(2)
                # **使用 JavaScript 移除 `aria-hidden="true"`**
                try:
                    driver.execute_script('document.querySelector("#modal_new_ec").style.display = "none";')
                    log_message("✅ 成功關閉 `修模通知表`，確保彈窗可見")
                except:
                    log_message("⚠️ 找不到 `修模通知表`，可能已經可見")
                try:
                    # **點擊 `x` 關閉視窗**
                    close_button = driver.find_element(By.XPATH, "//button[@aria-label='Close']")
                    driver.execute_script("arguments[0].click();", close_button)
                    log_message("✅ 使用 JavaScript 成功點擊 `x` 關閉視窗")
                except NoSuchElementException:
                    log_message("⚠️ 找不到 `x` 按鈕，跳過")
                time.sleep(1)
                # **處理 iframe 問題**
                try:
                    iframe = driver.find_element(By.TAG_NAME, "iframe")
                    driver.switch_to.frame(iframe)
                    print("🔄 切換至 iframe")
                except NoSuchElementException:
                    print("⚠️ 無需切換 iframe")

                try:
                    select_element = wait.until(EC.presence_of_element_located((By.ID, "wfn04")))
                    driver.execute_script("arguments[0].scrollIntoView();", select_element)
                    Select(select_element).select_by_visible_text("同意")
                    log_message("✅ 下拉選單已選擇")
                except TimeoutException:
                    log_message("❌ 找不到下拉選單")

                try:
                    text_area = wait.until(EC.presence_of_element_located((By.ID, "wfn10")))
                    text_area.send_keys(input_text)
                    log_message("✅ 文字已輸入")
                except TimeoutException:
                    log_message("❌ 找不到輸入框")

                # **使用 JavaScript 點擊送出按鈕**
                try:
                    submit_button = wait.until(EC.presence_of_element_located((By.ID, "save")))
                    driver.execute_script("arguments[0].click();", submit_button)
                    print("✅ 使用 JavaScript 成功點擊送出按鈕！")
                except TimeoutException:
                    print("❌ 送出按鈕無法點擊，請檢查是否有遮擋元素。")

                # **處理彈出視窗**
                time.sleep(1)
                try:
                    alert = driver.switch_to.alert
                    print(f"⚠️ 偵測到彈出視窗: {alert.text}，正在點擊確定...")
                    alert.accept()  # 點擊「確定」
                    print("✅ 已成功關閉彈出視窗")
                except NoAlertPresentException:
                    print("✅ 沒有偵測到彈出視窗，繼續執行")

                # **返回「我的簽呈」頁面**
                driver.back()
                time.sleep(1)
                log_message("🔄 返回簽呈列表頁面，準備處理下一個")

            except UnexpectedAlertPresentException:
                log_message("⚠️ 偵測到彈出視窗，強制關閉")
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                    log_message("✅ 已成功關閉彈出視窗")
                except NoAlertPresentException:
                    log_message("⚠️ 彈出視窗已被關閉，無需處理")

            except StaleElementReferenceException:
                log_message("⚠️ 簽呈已變更，重新取得資料")
                break

    log_message("✅ 所有簽核處理完成")

root.mainloop()

#pyinstaller --onefile --noconsole --icon="002.ico" --add-data "002.ico;." --add-data "chrome.ico;." --add-data "edge.ico;." --add-binary "C:/chromedriver/chromedriver.exe;./" --add-binary "C:/edgedriver/msedgedriver.exe;./" DMS_V2.0.py