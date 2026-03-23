# NavisworksMCP

AI 驅動的 Navisworks 自動化控制系統
透過 **Model Context Protocol (MCP)** 讓 Claude、Gemini、VS Code Copilot 等 AI 平台直接操控 Autodesk Navisworks。

---

## 系統架構

```
┌──────────────┐      stdio       ┌──────────────┐    WebSocket     ┌──────────────────┐
│  AI 平台      │ ◄──────────────► │  MCP Server  │ ◄──────────────► │  Navisworks 插件  │
│  Claude       │     MCP 協議     │  (Node.js)   │    port 2233    │  (C# Add-in)     │
│  Gemini       │                  │              │                  │                  │
│  VS Code      │                  │  index.ts    │                  │  Application.cs  │
└──────────────┘                  └──────────────┘                  └──────────────────┘
```

**三層分離設計：**
- **AI 平台** — 發送自然語言指令，由 MCP 協議轉譯為工具呼叫
- **MCP Server** — Node.js/TypeScript 中繼層，管理 WebSocket 連線與命令路由
- **Navisworks Add-in** — C# Plugin，在 NW 主執行緒中安全執行 .NET API 操作

---

## 快速開始

### 環境需求

| 項目 | 版本 |
|------|------|
| Navisworks Manage | 2025 |
| .NET Framework | 4.8 |
| Node.js | 20+ LTS |
| Visual Studio | 2019+ |

### 安裝步驟

```bash
# 1. 一鍵安裝 MCP Server 依賴並編譯
scripts\setup.bat

# 2. 在 Visual Studio 開啟解決方案
#    → 確認 NavisworksMCP.csproj 中的 NavisworksPath 正確
#    → Build Release

# 3. 部署 Plugin（需要管理員權限）
powershell -ExecutionPolicy Bypass -File scripts\install-addon.ps1
```

### 啟動服務

1. 開啟 Navisworks Manage 2025
2. 載入你的 NWF/NWD 模型
3. 在 **Add-ins** 面板找到 **"MCP 服務 (開/關)"** 按鈕，點擊啟動
4. 看到提示 `WebSocket 端口: 2233` 即表示服務就緒

### 連接 AI 平台

**VS Code / Claude Code：** 已配置 `.vscode/mcp.json`，開箱即用

**Claude Desktop：** 在設定中加入：
```json
{
  "mcpServers": {
    "navisworks": {
      "command": "node",
      "args": ["C:/Users/你的帳號/DESKTOP/NV_mcp/MCP-Server/build/index.js"],
      "env": { "NAVIS_MCP_PORT": "2233" }
    }
  }
}
```

---

## 使用模式

NavisworksMCP 提供 **27 個工具**，設計為五步驟工作流程：

### Step 1 — 了解模型

> 「這個模型有什麼？裡面有幾個來源檔案？主要分類是哪些？」

```
get_document_info          → 文件名稱、單位、模型數量
get_model_info             → 各來源檔案的根節點與子項目數
get_model_statistics       → 按分類/圖層/來源檔案統計數量  ★ 核心工具
get_all_categories         → 列出所有屬性分類名稱
get_model_tree             → 展開模型樹結構（可指定深度）
```

**典型對話：**
```
用戶：「幫我分析一下這個模型的組成」
AI 呼叫：get_model_statistics(groupBy: "sourceFile")
         → 回傳：Architectural.nwc (3,200 items), MEP.nwc (1,800 items), Structure.nwc (950 items)
AI 呼叫：get_model_statistics(groupBy: "classDisplayName")
         → 回傳：Wall (420), Door (180), Pipe (650), Duct (320)...
```

---

### Step 2 — 篩選與定位

> 「找出所有的門」「選擇 Selection Set 裡的機電管線」「用搜尋集過濾」

```
get_selection_sets         → 列出所有已儲存的選擇集
get_selection_set_items    → 讀取 Set 內容（不改變選擇）   ★ 安全讀取
execute_search_set         → 執行 Search Set 並回傳結果    ★ 搜尋集
search_items               → 自定義搜尋（按分類+屬性+值）
select_items_by_search     → 搜尋並選擇項目
select_items_by_set        → 選擇指定 Set 的所有項目
clear_selection            → 清除選擇
```

**典型對話：**
```
用戶：「模型裡有哪些選擇集？」
AI 呼叫：get_selection_sets
         → 回傳：MEP-Piping (Search Set, 650 items), Clash-Group-A (120 items)...

用戶：「看一下 MEP-Piping 裡面有什麼」
AI 呼叫：get_selection_set_items(name: "MEP-Piping", maxResults: 50)
         → 回傳項目清單（不影響當前選擇狀態）

用戶：「找出所有 Element 分類中 Type 是 Wall-200mm 的項目」
AI 呼叫：search_items(category: "Element", property: "Type", value: "Wall-200mm")
```

**`get_selection_set_items` vs `select_items_by_set` 的差別：**

| 工具 | 改變選擇？ | 用途 |
|------|-----------|------|
| `get_selection_set_items` | 不改變 | 安全探索，了解 Set 內容 |
| `select_items_by_set` | 改變 | 需要後續操作（上色、隱藏等）時使用 |

---

### Step 3 — 偵測現有狀態

> 「哪些東西被隱藏了？有沒有被凍結的項目？模型的篩選狀態是什麼？」

```
get_override_status        → 綜合報告：隱藏/凍結概覽       ★ 狀態總覽
get_hidden_items           → 列出所有被隱藏的項目
get_frozen_items           → 列出所有被凍結的項目
```

**典型對話：**
```
用戶：「看看模型現在的篩選狀態」
AI 呼叫：get_override_status(scope: "all", maxResults: 50)
         → 回傳：隱藏 23 項、凍結 5 項，附帶路徑清單

用戶：「哪些管線被隱藏了？」
AI 呼叫：get_hidden_items(maxResults: 200)
         → 回傳所有隱藏項目的名稱、分類、路徑
```

---

### Step 4 — 資料抽取

> 「把所有門的 ID 和高度匯出」「讀取選擇集的特定屬性」

```
get_item_properties        → 單項完整屬性（所有分類）
batch_get_properties       → 批量抽取，可指定欄位          ★ 資料匯出
get_item_geometry_info     → 幾何包圍盒與中心點座標
```

**典型對話：**
```
用戶：「把 MEP-Piping 選擇集裡所有管線的直徑和材質抽出來」
AI 呼叫：batch_get_properties(
           setName: "MEP-Piping",
           fields: ["Element.Diameter", "Element.Material"],
           maxResults: 500
         )
         → 回傳表格式資料，每列包含 displayName、path、指定欄位值

用戶：「看一下選中項目的完整屬性」
AI 呼叫：get_item_properties
         → 回傳所有分類 > 屬性名 > 值 > 資料類型
```

**`batch_get_properties` 兩種模式：**

| 模式 | 參數 | 說明 |
|------|------|------|
| 從 Selection Set | `setName: "MEP-Piping"` | 直接讀取，不影響選擇 |
| 從當前選擇 | 不指定 `setName` | 讀取 `CurrentSelection` |

---

### Step 5 — 視覺管理

> 「把管線標紅色」「隱藏結構柱」「切換到 3D 視點」

```
set_item_override_color    → 覆蓋選中項目的顯示顏色 (RGB)
clear_override_colors      → 清除顏色覆蓋（重設材質）
hide_items                 → 隱藏選中項目
unhide_all                 → 取消隱藏所有項目
zoom_to_selection          → 視圖對準選中項目
get_viewpoints             → 列出已儲存的視點
set_active_viewpoint       → 切換到指定視點
get_current_selection      → 確認當前選了什麼
```

**典型對話：**
```
用戶：「把碰撞區域的管線標成紅色」
AI 呼叫：select_items_by_set(name: "Clash-Group-A")
AI 呼叫：set_item_override_color(r: 255, g: 0, b: 0)
AI 呼叫：zoom_to_selection

用戶：「隱藏所有結構柱，只看機電」
AI 呼叫：search_items(category: "Element", property: "Category", value: "Structural Columns")
AI 呼叫：select_items_by_search(category: "Element", property: "Category", value: "Structural Columns")
AI 呼叫：hide_items

用戶：「恢復全部顯示」
AI 呼叫：unhide_all
AI 呼叫：clear_override_colors
```

---

## 完整工具清單

### 讀取工具（不改變模型狀態）— 17 個

| 工具 | 說明 |
|------|------|
| `get_document_info` | 文件基本資訊 |
| `get_model_info` | 來源模型清單 |
| `get_model_tree` | 模型樹結構 |
| `get_model_statistics` | 按分類/圖層/來源統計 |
| `get_all_categories` | 屬性分類名稱 |
| `get_current_selection` | 當前選擇清單 |
| `get_item_properties` | 單項屬性詳情 |
| `get_item_geometry_info` | 幾何包圍盒 |
| `get_selection_sets` | 選擇集清單 |
| `get_selection_set_items` | 選擇集內容（不選擇） |
| `get_viewpoints` | 視點清單 |
| `get_clash_tests` | 碰撞測試清單 |
| `get_clash_results` | 碰撞結果詳情 |
| `get_override_status` | 隱藏/凍結狀態概覽 |
| `get_hidden_items` | 隱藏項目清單 |
| `get_frozen_items` | 凍結項目清單 |
| `batch_get_properties` | 批量屬性抽取 |

### 操作工具（會改變模型狀態）— 10 個

| 工具 | 改變什麼 |
|------|----------|
| `search_items` | 不改變選擇，僅回傳結果 |
| `select_items_by_search` | 改變 CurrentSelection |
| `select_items_by_set` | 改變 CurrentSelection |
| `execute_search_set` | 可選擇是否改變 CurrentSelection |
| `clear_selection` | 清空 CurrentSelection |
| `set_active_viewpoint` | 切換視點 |
| `set_item_override_color` | 覆蓋顏色 |
| `clear_override_colors` | 重設材質 |
| `hide_items` | 隱藏項目 |
| `unhide_all` | 取消隱藏 |

---

## 工具組合範例

### 範例 1：模型初始分析

```
1. get_document_info                          → 了解基本資訊
2. get_model_statistics(groupBy: "sourceFile") → 看各來源檔案佔比
3. get_model_statistics(groupBy: "classDisplayName") → 看分類分佈
4. get_selection_sets                          → 看有哪些預設篩選
5. get_override_status(scope: "all")           → 看是否有已套用的篩選
```

### 範例 2：管線資料匯出

```
1. get_selection_sets                          → 找到 "MEP-Piping" 集
2. get_selection_set_items(name: "MEP-Piping") → 預覽內容
3. batch_get_properties(                       → 批量抽取指定欄位
     setName: "MEP-Piping",
     fields: ["Element.Id", "Element.Diameter", "Element.Material"]
   )
```

### 範例 3：碰撞項目視覺化

```
1. get_selection_sets                          → 找碰撞相關 Set
2. select_items_by_set(name: "Clash-Group-A")  → 選擇碰撞群組
3. set_item_override_color(r:255, g:0, b:0)    → 標紅色
4. zoom_to_selection                           → 對準畫面
5. get_item_properties                         → 看碰撞項目屬性
```

### 範例 4：隱藏非相關學科

```
1. get_model_statistics(groupBy: "sourceFile") → 確認有哪些來源
2. search_items(category: "Item", property: "Source File", value: "Structure.nwc")
3. select_items_by_search(同上參數)             → 選擇結構模型
4. hide_items                                  → 隱藏結構
5. unhide_all                                  → 完成後恢復
```

---

## 專案結構

```
NV_mcp/
├── NavisworksMCP.sln           # Visual Studio 解決方案
├── CLAUDE.md                   # AI 開發規範
├── README.md                   # 本文件
├── .gitignore
│
├── MCP/                        # C# Navisworks Add-in
│   ├── NavisworksMCP.csproj    # .NET Framework 4.8
│   ├── Application.cs          # Plugin 入口 (Toggle + Log + EventWatcher)
│   ├── Core/
│   │   ├── SocketService.cs    # WebSocket 服務 (port 2233)
│   │   ├── CommandExecutor.cs  # 27 個命令的執行器
│   │   ├── IdleEventManager.cs # 主執行緒安全排程
│   │   └── Logger.cs           # 日誌系統
│   └── Models/
│       ├── NavisCommandRequest.cs
│       └── NavisCommandResponse.cs
│
├── MCP-Server/                 # Node.js MCP Server
│   ├── package.json            # @modelcontextprotocol/sdk + ws
│   ├── tsconfig.json
│   └── src/
│       ├── index.ts            # MCP Server 入口 (stdio transport)
│       ├── socket.ts           # WebSocket 客戶端
│       └── tools/
│           └── navis-tools.ts  # 27 個工具定義
│
├── scripts/
│   ├── setup.bat               # 一鍵安裝
│   └── install-addon.ps1       # Plugin 部署
│
└── .vscode/
    └── mcp.json                # VS Code MCP 配置
```

---

## 注意事項

- WebSocket 端口固定為 **2233**，可透過環境變數 `NAVIS_MCP_PORT` 調整
- 所有命令在 Navisworks **主執行緒**中序列化執行，不會有並發衝突
- Plugin DLL 部署到 `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins\NavisworksMCP\`
- 日誌存放在 `%APPDATA%\NavisworksMCP\Logs\`

## 授權

MIT License
