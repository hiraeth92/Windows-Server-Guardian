# 🛡️ Windows Server Guardian

📌 專案描述（Project Description）

**中文簡介：**  
Windows Server Guardian 是一款專為 Windows 系統設計的自動化監控平台，可即時追蹤 CPU、記憶體、磁碟、GPU 使用率與網路流量，並自動掃描系統錯誤與警告日誌，當偵測到異常情況時，會以 Email 發出警告通知。適合作為系統穩定性監控、IT 管理與自動化警報應用。

**English Description:**  
Windows Server Guardian is an automated monitoring platform designed for Windows systems. It tracks CPU, memory, disk, GPU usage, and network traffic in real-time. It also scans Windows event logs for errors and warnings. When abnormal conditions are detected, it sends alert emails automatically. Ideal for system reliability monitoring, IT management, and automated alerting solutions.

---

## ✅ Features

- 📊 Real-time monitoring of:
  - CPU usage
  - Memory usage
  - Disk usage
  - GPU usage
  - Network I/O (upload/download speed)
- 📁 Automatic scanning of Windows event logs
- 📧 Email alert system for abnormal conditions
- 🧠 Keyword filtering and threshold-based warnings
- ⏱️ Fully automated: checks every 10 seconds and alerts every 5 minutes if needed

---

## 📦 Installation

```bash
pip install -r requirements.txt
