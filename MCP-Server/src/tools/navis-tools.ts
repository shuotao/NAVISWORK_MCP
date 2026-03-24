import { NavisSocketClient } from "../socket.js";

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, unknown>;
    required?: string[];
  };
}

/** 定義所有 Navisworks MCP 工具 */
export function getNavisTools(): ToolDefinition[] {
  return [
    // ─── 文件與模型資訊 ───
    {
      name: "get_document_info",
      description:
        "取得目前開啟的 Navisworks 文件資訊（標題、檔名、單位、模型數量）",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "get_model_info",
      description: "取得所有已載入模型的詳細資訊（檔名、根節點、子項目數量）",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "get_model_tree",
      description: "取得模型樹結構，可指定展開深度",
      inputSchema: {
        type: "object",
        properties: {
          maxDepth: {
            type: "number",
            description: "最大展開深度（預設 3）",
          },
        },
      },
    },

    // ─── 選擇操作 ───
    {
      name: "get_current_selection",
      description: "取得目前選擇的項目清單",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "select_items_by_search",
      description:
        "透過屬性搜尋並選擇項目（例如按 Element ID、名稱等）",
      inputSchema: {
        type: "object",
        properties: {
          category: {
            type: "string",
            description: "屬性分類名稱（例如 'Element'、'Item'）",
          },
          property: {
            type: "string",
            description: "屬性名稱（例如 'Id'、'Name'）",
          },
          value: {
            type: "string",
            description: "要搜尋的值",
          },
        },
        required: ["category", "property", "value"],
      },
    },
    {
      name: "clear_selection",
      description: "清除目前的選擇",
      inputSchema: { type: "object", properties: {} },
    },

    // ─── 屬性查詢 ───
    {
      name: "get_item_properties",
      description:
        "取得選中項目或搜尋項目的所有屬性（按分類和屬性名列出）",
      inputSchema: {
        type: "object",
        properties: {
          category: { type: "string", description: "搜尋分類（可選）" },
          property: { type: "string", description: "搜尋屬性（可選）" },
          value: { type: "string", description: "搜尋值（可選）" },
        },
      },
    },
    {
      name: "get_all_categories",
      description: "取得模型中所有可用的屬性分類名稱",
      inputSchema: { type: "object", properties: {} },
    },

    // ─── 搜尋 ───
    {
      name: "search_items",
      description:
        "在模型中搜尋項目，支援多種比對條件（equals、contains、greater_than、less_than）",
      inputSchema: {
        type: "object",
        properties: {
          category: {
            type: "string",
            description: "屬性分類名稱",
          },
          property: {
            type: "string",
            description: "屬性名稱（可選，若只指定 category 則搜尋該分類下所有項目）",
          },
          value: { type: "string", description: "搜尋值" },
          condition: {
            type: "string",
            enum: [
              "equals",
              "contains",
              "not_equals",
              "greater_than",
              "less_than",
            ],
            description: "比對條件（預設 equals）",
          },
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 100）",
          },
        },
        required: ["category"],
      },
    },

    // ─── 視點 ───
    {
      name: "get_viewpoints",
      description: "取得所有已儲存的視點清單",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "set_active_viewpoint",
      description: "切換到指定名稱的已儲存視點",
      inputSchema: {
        type: "object",
        properties: {
          name: { type: "string", description: "視點名稱" },
        },
        required: ["name"],
      },
    },

    // ─── Clash Detection ───
    {
      name: "get_clash_tests",
      description: "取得所有碰撞檢測測試清單及狀態",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "get_clash_results",
      description: "取得指定碰撞測試的詳細結果",
      inputSchema: {
        type: "object",
        properties: {
          testName: {
            type: "string",
            description: "碰撞測試名稱",
          },
        },
        required: ["testName"],
      },
    },

    // ─── 選擇集 ───
    {
      name: "get_selection_sets",
      description: "取得所有已儲存的選擇集",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "select_items_by_set",
      description: "選擇指定選擇集中的所有項目",
      inputSchema: {
        type: "object",
        properties: {
          name: { type: "string", description: "選擇集名稱" },
        },
        required: ["name"],
      },
    },

    // ─── 幾何與顯示 ───
    {
      name: "get_item_geometry_info",
      description: "取得選中項目的幾何資訊（包圍盒、中心點）",
      inputSchema: {
        type: "object",
        properties: {
          category: { type: "string", description: "搜尋分類（可選）" },
          property: { type: "string", description: "搜尋屬性（可選）" },
          value: { type: "string", description: "搜尋值（可選）" },
        },
      },
    },
    {
      name: "zoom_to_selection",
      description: "將視圖縮放至目前選擇的項目",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "set_item_override_color",
      description: "覆蓋選中項目的顯示顏色 (RGB 0-255)",
      inputSchema: {
        type: "object",
        properties: {
          r: { type: "number", description: "紅色 (0-255)" },
          g: { type: "number", description: "綠色 (0-255)" },
          b: { type: "number", description: "藍色 (0-255)" },
        },
        required: ["r", "g", "b"],
      },
    },
    {
      name: "clear_override_colors",
      description: "清除選中項目（或所有項目）的顏色覆蓋",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "hide_items",
      description: "隱藏目前選擇的項目",
      inputSchema: { type: "object", properties: {} },
    },
    {
      name: "unhide_all",
      description: "取消隱藏所有項目",
      inputSchema: { type: "object", properties: {} },
    },

    // ─── 篩選與資料抽取（新增） ───
    {
      name: "get_selection_set_items",
      description:
        "取得指定 Selection Set 內的所有項目（不改變選擇狀態），包含 Search Set 自動執行",
      inputSchema: {
        type: "object",
        properties: {
          name: { type: "string", description: "選擇集名稱" },
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 200）",
          },
        },
        required: ["name"],
      },
    },
    {
      name: "execute_search_set",
      description:
        "執行已儲存的 Search Set（搜尋集），可選擇是否同時選擇結果",
      inputSchema: {
        type: "object",
        properties: {
          name: { type: "string", description: "搜尋集名稱" },
          select: {
            type: "boolean",
            description: "是否同時選擇搜尋結果（預設 true）",
          },
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 200）",
          },
        },
        required: ["name"],
      },
    },
    {
      name: "get_override_status",
      description:
        "偵測模型中被隱藏、凍結的項目狀態概覽（用於了解模型當前的篩選狀態）",
      inputSchema: {
        type: "object",
        properties: {
          scope: {
            type: "string",
            enum: ["selection", "all"],
            description:
              "'selection' 只掃描已選擇項目，'all' 掃描整個模型（預設 selection）",
          },
          maxResults: {
            type: "number",
            description: "每類最大回傳數量（預設 200）",
          },
        },
      },
    },
    {
      name: "get_hidden_items",
      description: "列出模型中所有被隱藏的項目",
      inputSchema: {
        type: "object",
        properties: {
          scope: {
            type: "string",
            enum: ["selection", "all"],
            description: "掃描範圍（預設 all）",
          },
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 500）",
          },
        },
      },
    },
    {
      name: "get_frozen_items",
      description: "列出模型中所有被凍結的項目",
      inputSchema: {
        type: "object",
        properties: {
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 500）",
          },
        },
      },
    },
    {
      name: "batch_get_properties",
      description:
        "批量抽取多個項目的屬性資料（從當前選擇或指定 Selection Set），可指定要抽取的欄位",
      inputSchema: {
        type: "object",
        properties: {
          setName: {
            type: "string",
            description:
              "從指定 Selection Set 抽取（可選，若不指定則從當前選擇抽取）",
          },
          fields: {
            type: "array",
            items: { type: "string" },
            description:
              "要抽取的欄位陣列，格式 'Category.Property'（例如 ['Element.Id', 'Item.Layer']）。若不指定則抽取所有屬性",
          },
          maxResults: {
            type: "number",
            description: "最大回傳數量（預設 500）",
          },
        },
      },
    },
    {
      name: "get_model_statistics",
      description:
        "模型統計分析 — 按分類/圖層/來源檔案彙總項目數量，快速了解模型組成",
      inputSchema: {
        type: "object",
        properties: {
          groupBy: {
            type: "string",
            enum: ["classDisplayName", "className", "layer", "sourceFile"],
            description:
              "分組方式：classDisplayName（顯示分類）、className（內部分類）、layer（圖層）、sourceFile（來源檔案）",
          },
          geometryOnly: {
            type: "boolean",
            description: "是否只統計有幾何的項目（預設 true）",
          },
        },
      },
    },
    // ─── 子樹掃描 ───
    {
      name: "scan_subtree",
      description:
        "掃描指定 NWC/NWD 子樹節點下的所有幾何元素，按 Revit Category 和 Family/Type 分組統計。用於精確掃描特定來源檔案的內容，不受全模型取樣限制。",
      inputSchema: {
        type: "object",
        properties: {
          name: {
            type: "string",
            description:
              "要掃描的節點名稱（例如 'DE_MSI_001_CUB_B001_A_ARCH.nwc'），支援部分匹配",
          },
          maxItems: {
            type: "number",
            description: "最大掃描後代數量（預設 50000）",
          },
          fields: {
            type: "array",
            items: { type: "string" },
            description:
              "額外要讀取的屬性欄位（例如 ['Element.System Name', 'Element.Size']），指定後會回傳 rows 明細",
          },
          maxResults: {
            type: "number",
            description: "明細行最大數量（預設 5000，需配合 fields 使用）",
          },
        },
        required: ["name"],
      },
    },
    {
      name: "select_subtree",
      description:
        "選取指定子樹節點下的所有幾何元素到 CurrentSelection，可搭配 batch_get_properties 使用",
      inputSchema: {
        type: "object",
        properties: {
          name: {
            type: "string",
            description: "要選取的節點名稱，支援部分匹配",
          },
        },
        required: ["name"],
      },
    },
  ];
}

/** 執行 Navisworks 工具 */
export async function executeNavisTool(
  client: NavisSocketClient,
  toolName: string,
  args: Record<string, unknown>
): Promise<string> {
  const response = await client.sendCommand(toolName, args);

  if (!response.success) {
    throw new Error(response.error || `命令執行失敗: ${toolName}`);
  }

  return JSON.stringify(response.data, null, 2);
}
