# 🛡️ Server Guardian

Windows 系統資源與錯誤日誌自動監控平台  
**Automated Monitoring Platform for Windows System Resources and Logs**

---

## 📘 專案簡介 | Project Overview

**Server Guardian** 是一款專為 Windows 系統設計的自動化監控工具，能即時監控 CPU、記憶體、磁碟、GPU 使用率與網路流量，並自動掃描系統錯誤與警告日誌。當偵測到異常情況時，會自動發送 Email 警示通知。  
本專案適用於系統穩定性監控、IT 管理與自動化警報應用。

**Server Guardian** is a real-time automated monitoring tool for Windows systems. It monitors CPU, memory, disk, GPU usage, and network traffic. It also scans system error and warning logs and sends alert emails automatically when issues are detected.  
Ideal for system reliability monitoring, IT operations, and automated alerting scenarios.

---

## ⚙️ 功能特色 | Features

- ✅ 實時監控 CPU / 記憶體 / 磁碟 / GPU 使用率  
- ✅ 網路流量監測（上傳/下載）  
- ✅ 每五分鐘自動檢查系統錯誤與警告日誌  
- ✅ Email 告警通知（自動寄送摘要）  
- ✅ 具備資源與關鍵字過濾機制  
- ✅ 使用 `.env` 檔案安全管理帳號密碼  

---

## 📦 安裝說明 | Installation

### 1️⃣ 下載專案 Clone the Repository

```bash
git clone https://github.com/你的帳號/Windows-Server-Guardian.git
cd Windows-Server-Guardian
