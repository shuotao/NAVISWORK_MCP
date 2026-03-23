# NavisworksMCP

AI 驅動的 Navisworks 自動化控制系統，透過 Model Context Protocol (MCP) 讓 Claude、Gemini 等 AI 平台操控 Autodesk Navisworks。

## 架構

```
AI 平台 (Claude/Gemini) ←→ MCP Server (Node.js, stdio) ←→ WebSocket (port 2222) ←→ Navisworks Add-in (C#)
```

### 三層通訊
1. **AI 平台** — 透過 MCP 協議發送工具調用
2. **MCP Server** (`MCP-Server/`) — Node.js/TypeScript，將 MCP 請求轉譯為 WebSocket 命令
3. **Navisworks Add-in** (`MCP/`) — C# Plugin，接收命令並透過 Navisworks .NET API 執行

## 開發規則

- WebSocket 端口固定為 **2222**
- C# 專案目標 .NET Framework 4.8
- Navisworks API 引用的 Copy Local 必須設為 False
- 命令在 Navisworks 主執行緒中透過 Idle 事件執行
- MCP Server 使用 stdio transport

## 專案結構

```
MCP/                    # C# Navisworks Add-in
├── Application.cs      # Plugin 入口 (AddInPlugin + EventWatcher)
├── Core/
│   ├── SocketService.cs      # WebSocket 服務 (port 2222)
│   ├── CommandExecutor.cs    # 命令執行器
│   ├── IdleEventManager.cs   # 主執行緒排程
│   └── Logger.cs             # 日誌
├── Models/             # 資料模型
└── NavisworksMCP.csproj

MCP-Server/             # Node.js MCP Server
├── src/
│   ├── index.ts        # MCP Server 入口
│   ├── socket.ts       # WebSocket 客戶端
│   └── tools/
│       └── navis-tools.ts  # 工具定義
├── package.json
└── tsconfig.json
```

## 可用工具

### 文件與模型
- `get_document_info` / `get_model_info` / `get_model_tree`

### 選擇與搜尋
- `get_current_selection` / `select_items_by_search` / `clear_selection`
- `search_items` / `get_all_categories`

### 屬性
- `get_item_properties` / `get_item_geometry_info`

### 視點
- `get_viewpoints` / `set_active_viewpoint`

### Clash Detection
- `get_clash_tests` / `get_clash_results`

### 選擇集
- `get_selection_sets` / `select_items_by_set`

### 顯示控制
- `zoom_to_selection` / `set_item_override_color` / `clear_override_colors`
- `hide_items` / `unhide_all`
