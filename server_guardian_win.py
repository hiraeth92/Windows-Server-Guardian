import os
import psutil
import GPUtil
import time
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from dotenv import load_dotenv
import win32evtlog
import re

load_dotenv()  # 載入 .env 環境變數

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")      # 寄件郵箱帳號
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")    # 寄件郵箱密碼
TO_EMAIL = "自己的帳號@gmail.com"                 # 收件郵箱地址

# 系統資源使用率警戒門檻設定
CPU_THRESHOLD = 80
MEM_THRESHOLD = 80
DISK_THRESHOLD = 90
GPU_THRESHOLD = 80

EMAIL_INTERVAL = 300  # 郵件寄送間隔秒數(5分鐘)

# 日誌過濾關鍵字，跳過含這些詞的訊息 (忽略大小寫，用正則表示)
SKIP_KEYWORDS = [re.compile(k, re.I) for k in ["DCOM", "DNS Client"]]
# 日誌警告關鍵字，只要訊息含有就會視為警告 (忽略大小寫，用正則表示)
ALERT_KEYWORDS = [re.compile(k, re.I) for k in ["error", "fail", "critical"]]

def send_email(subject, body, sender, password, receiver):
    """
    使用 Gmail SMTP SSL 發送 Email
    subject: 郵件主旨
    body: 郵件內容
    sender: 寄件郵箱帳號
    password: 寄件郵箱密碼
    receiver: 收件人地址
    """
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver
    try:
        # 連線 Gmail SMTP SSL，登入並發送郵件
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)
        print("✅ 已發送 Email 警告")
    except Exception as e:
        print(f"Email 發送失敗: {e}")

def check_windows_logs(server='localhost', minutes=5):
    """
    讀取 Windows 系統日誌
    回傳最近 N 分鐘內符合警告條件的訊息列表
    過濾規則：訊息需包含 ALERT_KEYWORDS 中任一字串，且不包含 SKIP_KEYWORDS 任何字串
    """
    log_type = 'System'  # 系統事件日誌類型
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    hand = win32evtlog.OpenEventLog(server, log_type)
    alerts = []
    now = datetime.now()
    time_limit = now - timedelta(minutes=minutes)  # 時間範圍起點

    while True:
        events = win32evtlog.ReadEventLog(hand, flags, 0)
        if not events:
            break
        for event in events:
            try:
                # 將事件時間字串轉成 datetime 物件，格式 MM/DD/YYYY HH:MM:SS
                event_time = datetime.strptime(event.TimeGenerated.Format(), '%m/%d/%Y %H:%M:%S')
            except Exception:
                continue
            if event_time < time_limit:
                # 事件時間早於查詢範圍，結束查詢並回傳結果
                win32evtlog.CloseEventLog(hand)
                return alerts
            # 僅判斷錯誤與警告類型事件
            if event.EventType in [win32evtlog.EVENTLOG_ERROR_TYPE, win32evtlog.EVENTLOG_WARNING_TYPE]:
                source = str(event.SourceName)
                try:
                    msg = win32evtlog.FormatMessage(event)  # 取得詳細訊息
                except:
                    msg = "無法取得訊息"
                full_msg = f"[{event_time.strftime('%Y-%m-%d %H:%M:%S')}] {source}: {msg.strip()}"
                # 判斷是否符合警告條件並排除不需處理的關鍵字
                if (any(p.search(full_msg) for p in ALERT_KEYWORDS) and
                    not any(p.search(full_msg) for p in SKIP_KEYWORDS)):
                    alerts.append(full_msg)

    win32evtlog.CloseEventLog(hand)
    return alerts

def get_system_status(prev_net):
    """
    取得系統資源使用狀態與網路流量變化
    prev_net: 前一次網路流量計數 (psutil.net_io_counters() 回傳物件)
    回傳：
    cpu, mem, disk, gpu 使用率百分比，
    最新網路計數物件 net_io，
    以及計算出的網路上傳/下載速度 KB/s
    """
    cpu = psutil.cpu_percent(interval=0)  # 非阻塞取得 CPU 使用率
    mem = psutil.virtual_memory().percent
    disk = psutil.disk_usage('C:\\').percent
    net_io = psutil.net_io_counters()
    net_sent = (net_io.bytes_sent - prev_net.bytes_sent) / 1024  # KB/s
    net_recv = (net_io.bytes_recv - prev_net.bytes_recv) / 1024  # KB/s
    gpus = GPUtil.getGPUs()
    gpu_usage = max([gpu.load * 100 for gpu in gpus]) if gpus else 0
    return cpu, mem, disk, gpu_usage, net_io, net_sent, net_recv

def monitor():
    """
    主監控迴圈
    每10秒更新系統資源狀態、讀取系統日誌
    5分鐘內若有警告錯誤則合併寄送 Email 警報
    """
    prev_net = psutil.net_io_counters()
    last_sent_time = 0
    alert_cache = []

    while True:
        cpu, mem, disk, gpu, net_io, net_sent, net_recv = get_system_status(prev_net)
        prev_net = net_io

        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        status = (f"[{timestamp}] CPU: {cpu:.1f}% | MEM: {mem:.1f}% | DISK: {disk:.1f}% | "
                  f"GPU: {gpu:.1f}% | NET Sent: {net_sent:.1f} KB/s | NET Recv: {net_recv:.1f} KB/s")
        print(status)

        alerts = []
        # 判斷各項資源是否超過設定門檻
        if cpu > CPU_THRESHOLD:
            alerts.append(f"CPU 使用率過高：{cpu:.1f}%")
        if mem > MEM_THRESHOLD:
            alerts.append(f"記憶體使用率過高：{mem:.1f}%")
        if disk > DISK_THRESHOLD:
            alerts.append(f"磁碟使用率過高：{disk:.1f}%")
        if gpu > GPU_THRESHOLD:
            alerts.append(f"GPU 使用率過高：{gpu:.1f}%")

        # 讀取近5分鐘系統日誌錯誤警告
        log_alerts = check_windows_logs(minutes=5)
        if log_alerts:
            alerts.append("系統日誌錯誤或警告：\n" + "\n".join(log_alerts))

        # 若有警告，印出並加入快取等待寄信
        if alerts:
            print("⚠️ 發現警告／錯誤：")
            for a in alerts:
                print(a)
            alert_cache.extend(alerts)

        # 超過EMAIL_INTERVAL秒且快取不為空，寄送 Email 並清空快取
        if time.time() - last_sent_time > EMAIL_INTERVAL and alert_cache:
            alert_text = "\n".join(alert_cache)
            send_email("📝 系統狀態摘要", alert_text, EMAIL_ADDRESS, EMAIL_PASSWORD, TO_EMAIL)
            alert_cache.clear()
            last_sent_time = time.time()

        time.sleep(10)  # 每10秒更新一次

if __name__ == "__main__":
    try:
        monitor()
    except KeyboardInterrupt:
        print("程式中斷，結束監控")
