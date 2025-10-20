# 通道檢查工具 (Channel Check Tool)

通道檢查工具是一款用於模擬和分析高速電子通道的桌面應用程式。它提供圖形化使用者介面 (GUI)，引導使用者完成匯入電路板設計、設定模擬、執行模擬以及分析信號完整性結果的整個流程。

## 功能

-   從 `.brd` 或 `.aedb` 檔案匯入電路板設計。
-   自動從設計中提取元件和網路資訊。
-   為控制器 (發送器) 和 DRAM (接收器) 元件設定埠。
-   使用 Ansys SIwave 設定並執行頻域模擬。
-   對產生的通道模型執行暫態模擬。
-   分析信號完整性指標，包括信號、符際干擾 (ISI) 和串擾。
-   查看並匯出模擬結果。

## 工作流程

本應用程式的介面由一系列頁籤組成，每個頁籤對應工作流程中的一個步驟：

1.  **匯入 (Import):** 首先匯入您的電路板設計。工具將提取必要的佈局資訊。
2.  **埠設定 (Port Setup):** 定義哪些元件是控制器，哪些是 DRAM。選擇您要分析的信號網路。此步驟會建立一個 `ports.json` 檔案來描述埠的設定。
3.  **模擬 (Simulation):** 設定並執行使用 SIwave 的頻域模擬。這將生成一個描述通道特性的 Touchstone (`.sNp`) 檔案。
4.  **CCT:** 載入 Touchstone 檔案和 `ports.json` 檔案。設定發送器 (TX) 和接收器 (RX) 的特性，並執行暫態模擬。
5.  **結果 (Result):** 查看 CCT 分析的結果，包括信號、ISI 和串擾等指標。

## 開始使用

### 先決條件

-   Python 3.x
-   Ansys Electronics Desktop (AEDT) with SIwave

### 安裝

1.  複製此儲存庫。
2.  安裝所需的 Python 套件：
    ```
    pip install -r requirements.txt
    ```

### 執行應用程式

若要啟動 GUI，請執行 `main.py` 指令碼：

```
python src/main.py
```

## 專案結構

-   `src/main.py`: 應用程式的主要進入點。
-   `src/gui.py`: 使用 PySide6 定義使用者介面。
-   `src/get_edb.py`: 從電路板設計中提取佈局資訊。
-   `src/set_edb.py`: 在 Ansys EDB 中設定埠。
-   `src/set_sim.py`: 設定 SIwave 模擬。
-   `src/run_sim.py`: 執行 SIwave 模擬。
-   `src/cct.py`: 通道檢查工具暫態模擬與分析的核心邏輯。
-   `src/cct_runner.py`: 從 GUI 執行 CCT 邏輯的輔助指令碼。
-   `data/`: 用於儲存資料檔案的目錄，例如 `ports.json`。
-   `requirements.txt`: 專案所需的 Python 套件列表。