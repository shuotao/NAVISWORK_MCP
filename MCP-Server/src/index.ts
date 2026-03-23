import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { NavisSocketClient } from "./socket.js";
import { getNavisTools, executeNavisTool } from "./tools/navis-tools.js";

const client = new NavisSocketClient();

const server = new Server(
  {
    name: "navisworks-mcp-server",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// 列出可用工具
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: getNavisTools() };
});

// 執行工具
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    const result = await executeNavisTool(
      client,
      name,
      (args as Record<string, unknown>) || {}
    );

    return {
      content: [{ type: "text", text: result }],
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      content: [{ type: "text", text: `錯誤: ${message}` }],
      isError: true,
    };
  }
});

// 啟動伺服器
async function main() {
  console.error("[NavisMCP] 正在啟動 MCP Server...");

  // 嘗試連接 Navisworks Add-in
  try {
    await client.connect();
    console.error("[NavisMCP] 已連接到 Navisworks Add-in");
  } catch {
    console.error(
      "[NavisMCP] 警告: 尚未連接到 Navisworks，將在收到命令時自動重連"
    );
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("[NavisMCP] MCP Server 已就緒");
}

// 優雅關閉
process.on("SIGINT", () => {
  client.disconnect();
  process.exit(0);
});
process.on("SIGTERM", () => {
  client.disconnect();
  process.exit(0);
});

main().catch((err) => {
  console.error("[NavisMCP] 啟動失敗:", err);
  process.exit(1);
});
