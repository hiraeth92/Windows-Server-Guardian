import os  # 匯入作業系統模組，用來存取環境變數等功能
import psutil  # 匯入系統資源監控模組，如 CPU、記憶體、磁碟、網路等
import GPUtil  # 匯入 GPU 使用率取得模組
import time  # 匯入時間模組，用來控制間隔與時間戳記
import smtplib  # 匯入 SMTP 郵件傳送模組
import threading  # 匯入多執行緒模組，用來同時處理不同任務
from email.mime.text import MIMEText  # 匯入 Email 文字內容格式模組
from datetime import datetime, timedelta  # 匯入時間與時間差物件
from dotenv import load_dotenv  # 匯入 .env 讀取模組
import win32evtlog  # 匯入 Windows 事件日誌讀取模組
import re  # 匯入正規表示式模組

load_dotenv()  # 載入 .env 環境變數檔案內容

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # 從環境變數中取得寄件者 Email
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # 從環境變數中取得寄件者密碼
TO_EMAIL = "自己的帳號@gmail.com"  # 收件者 Email

# 設定各項資源使用率的警戒門檻值
CPU_THRESHOLD = 80  # CPU 警戒值
MEM_THRESHOLD = 80  # 記憶體警戒值
DISK_THRESHOLD = 90  # 磁碟使用率警戒值
GPU_THRESHOLD = 80  # GPU 使用率警戒值

EMAIL_INTERVAL = 300  # 設定每封警告 Email 間的最小間隔（單位：秒）

# 設定日誌過濾條件
SKIP_KEYWORDS = [re.compile(k, re.I) for k in ["DCOM", "DNS Client"]]  # 忽略的日誌關鍵字（不理會大小寫）
ALERT_KEYWORDS = [re.compile(k, re.I) for k in ["error", "fail", "critical"]]  # 警告關鍵字（符合者將被記錄）

def send_email(subject, body, sender, password, receiver):
    """
    發送 Email 通知警告內容。
    """
    msg = MIMEText(body)  # 設定郵件主體文字內容
    msg['Subject'] = subject  # 設定郵件標題
    msg['From'] = sender  # 設定寄件者
    msg['To'] = receiver  # 設定收件者
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:  # 透過 Gmail SMTP 使用 SSL 傳送
            smtp.login(sender, password)  # 登入寄件帳號
            smtp.send_message(msg)  # 寄送郵件
        print("✅ 已發送 Email 警告")
    except Exception as e:
        print(f"Email 發送失敗: {e}")  # 發送失敗顯示錯誤訊息

def check_windows_logs(server='localhost', minutes=5):
    """
    檢查最近 N 分鐘內的 Windows 系統日誌，有無錯誤或警告訊息。
    傳回警告訊息列表。
    """
    log_type = 'System'  # 指定讀取的日誌類型為系統日誌
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ  # 從最新讀起，依序讀取
    hand = win32evtlog.OpenEventLog(server, log_type)  # 開啟日誌檔案
    alerts = []  # 儲存警告訊息列表
    now = datetime.now()
    time_limit = now - timedelta(minutes=minutes)  # 設定時間門檻（只讀取這段時間內的）

    while True:
        events = win32evtlog.ReadEventLog(hand, flags, 0)  # 讀取事件日誌
        if not events:
            break
        for event in events:
            try:
                event_time = datetime.strptime(event.TimeGenerated.Format(), '%m/%d/%Y %H:%M:%S')  # 轉換時間格式
            except Exception:
                continue
            if event_time < time_limit:
                win32evtlog.CloseEventLog(hand)  # 超過時間門檻就停止
                return alerts
            if event.EventType in [win32evtlog.EVENTLOG_ERROR_TYPE, win32evtlog.EVENTLOG_WARNING_TYPE]:
                source = str(event.SourceName)
                try:
                    msg = win32evtlog.FormatMessage(event)
                except:
                    msg = "無法取得訊息"
                full_msg = f"[{event_time.strftime('%Y-%m-%d %H:%M:%S')}] {source}: {msg.strip()}"
                if (any(p.search(full_msg) for p in ALERT_KEYWORDS) and
                    not any(p.search(full_msg) for p in SKIP_KEYWORDS)):
                    alerts.append(full_msg)  # 加入符合條件的警告訊息

    win32evtlog.CloseEventLog(hand)
    return alerts

def get_system_status(prev_net):
    """
    取得目前的系統資源狀態，並計算網路流量變化。
    傳回：CPU、記憶體、磁碟、GPU 使用率與目前網路資訊
    """
    cpu = psutil.cpu_percent(interval=0)  # CPU 使用率
    mem = psutil.virtual_memory().percent  # 記憶體使用率
    disk = psutil.disk_usage('C:\\').percent  # C 槽磁碟使用率
    net_io = psutil.net_io_counters()  # 目前網路傳輸總量
    net_sent = (net_io.bytes_sent - prev_net.bytes_sent) / 1024  # 上傳流量 KB/s
    net_recv = (net_io.bytes_recv - prev_net.bytes_recv) / 1024  # 下載流量 KB/s
    gpus = GPUtil.getGPUs()  # 取得 GPU 清單
    gpu_usage = max([gpu.load * 100 for gpu in gpus]) if gpus else 0  # 計算 GPU 使用率
    return cpu, mem, disk, gpu_usage, net_io, net_sent, net_recv

def monitor():
    """
    主監控邏輯：每 10 秒檢查系統狀態與日誌警告，觸發則發送 Email。
    """
    prev_net = psutil.net_io_counters()  # 初始網路流量數據
    last_sent_time = 0  # 上次寄出 Email 的時間戳
    alert_cache = []  # 警告訊息暫存

    while True:
        cpu, mem, disk, gpu, net_io, net_sent, net_recv = get_system_status(prev_net)  # 取得最新系統狀態
        prev_net = net_io  # 更新流量資料
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 取得目前時間
        status = (f"[{timestamp}] CPU: {cpu:.1f}% | MEM: {mem:.1f}% | DISK: {disk:.1f}% | "
                  f"GPU: {gpu:.1f}% | NET Sent: {net_sent:.1f} KB/s | NET Recv: {net_recv:.1f} KB/s")
        print(status)  # 輸出系統狀態

        alerts = []  # 本輪警告列表
        if cpu > CPU_THRESHOLD:
            alerts.append(f"CPU 使用率過高：{cpu:.1f}%")
        if mem > MEM_THRESHOLD:
            alerts.append(f"記憶體使用率過高：{mem:.1f}%")
        if disk > DISK_THRESHOLD:
            alerts.append(f"磁碟使用率過高：{disk:.1f}%")
        if gpu > GPU_THRESHOLD:
            alerts.append(f"GPU 使用率過高：{gpu:.1f}%")

        # 透過子執行緒讀取 Windows 日誌（不阻塞主執行緒）
        log_thread_result = []  # 儲存日誌結果的容器
        def read_log():
            log_thread_result.extend(check_windows_logs(minutes=5))  # 執行日誌讀取並存入結果
        log_thread = threading.Thread(target=read_log)  # 建立執行緒物件
        log_thread.start()  # 啟動執行緒
        log_thread.join()  # 等待子執行緒完成

        if log_thread_result:
            alerts.append("系統日誌錯誤或警告：\n" + "\n".join(log_thread_result))

        if alerts:
            print("⚠️ 發現警告／錯誤：")
            for a in alerts:
                print(a)  # 輸出警告訊息
            alert_cache.extend(alerts)  # 加入暫存

        if time.time() - last_sent_time > EMAIL_INTERVAL and alert_cache:
            alert_text = "\n".join(alert_cache)  # 組合警告文字
            email_thread = threading.Thread(  # 建立子執行緒寄信
                target=send_email,
                args=("📝 系統狀態摘要", alert_text, EMAIL_ADDRESS, EMAIL_PASSWORD, TO_EMAIL)
            )
            email_thread.start()  # 開始寄信執行緒
            last_sent_time = time.time()  # 更新上次寄信時間
            alert_cache.clear()  # 清除已寄出的警告內容

        time.sleep(10)  # 每 10 秒檢查一次

if __name__ == "__main__":
    try:
        monitor()  # 啟動監控主程式
    except KeyboardInterrupt:
        print("程式中斷，結束監控")  # 捕捉 Ctrl+C 結束程式
