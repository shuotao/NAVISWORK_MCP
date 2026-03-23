import WebSocket from "ws";

export interface NavisCommand {
  command: string;
  parameters: Record<string, unknown>;
  requestId: string;
}

export interface NavisResponse {
  success: boolean;
  data?: unknown;
  error?: string;
  requestId: string;
}

export class NavisSocketClient {
  private ws: WebSocket | null = null;
  private responseHandlers: Map<string, (response: NavisResponse) => void> =
    new Map();
  private reconnectInterval: NodeJS.Timeout | null = null;
  private readonly port: number;

  constructor() {
    this.port = parseInt(process.env.NAVIS_MCP_PORT || "2222", 10);
  }

  async connect(): Promise<void> {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error("連接超時 (10秒)"));
      }, 10000);

      try {
        this.ws = new WebSocket(`ws://localhost:${this.port}/`);

        this.ws.on("open", () => {
          clearTimeout(timeout);
          console.error(`[NavisMCP] 已連接到 Navisworks Add-in (port ${this.port})`);
          if (this.reconnectInterval) {
            clearInterval(this.reconnectInterval);
            this.reconnectInterval = null;
          }
          resolve();
        });

        this.ws.on("message", (data) => {
          try {
            const response: NavisResponse = JSON.parse(data.toString());
            const handler = this.responseHandlers.get(response.requestId);
            if (handler) {
              handler(response);
              this.responseHandlers.delete(response.requestId);
            }
          } catch (err) {
            console.error("[NavisMCP] 解析回應失敗:", err);
          }
        });

        this.ws.on("close", () => {
          console.error("[NavisMCP] 連接已關閉，嘗試重連...");
          this.startReconnect();
        });

        this.ws.on("error", (err) => {
          clearTimeout(timeout);
          console.error("[NavisMCP] WebSocket 錯誤:", err.message);
          reject(err);
        });
      } catch (err) {
        clearTimeout(timeout);
        reject(err);
      }
    });
  }

  private startReconnect(): void {
    if (this.reconnectInterval) return;
    this.reconnectInterval = setInterval(async () => {
      try {
        await this.connect();
      } catch {
        // 靜默重試
      }
    }, 5000);
  }

  async sendCommand(
    command: string,
    parameters: Record<string, unknown> = {}
  ): Promise<NavisResponse> {
    if (!this.isConnected()) {
      try {
        await this.connect();
      } catch {
        throw new Error(
          "無法連接到 Navisworks。請確認 Navisworks 已開啟且 MCP 服務已啟動 (port 2222)"
        );
      }
    }

    const requestId = this.generateRequestId();
    const cmd: NavisCommand = {
      command,
      parameters: this.toPascalCase(parameters),
      requestId,
    };

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.responseHandlers.delete(requestId);
        reject(new Error(`命令超時 (30秒): ${command}`));
      }, 30000);

      this.responseHandlers.set(requestId, (response) => {
        clearTimeout(timeout);
        resolve(response);
      });

      this.ws!.send(JSON.stringify(cmd));
    });
  }

  isConnected(): boolean {
    return this.ws?.readyState === WebSocket.OPEN;
  }

  disconnect(): void {
    if (this.reconnectInterval) {
      clearInterval(this.reconnectInterval);
      this.reconnectInterval = null;
    }
    if (this.ws) {
      this.ws.close();
      this.ws = null;
    }
  }

  private generateRequestId(): string {
    return `req_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }

  /** 將 camelCase key 轉換為 PascalCase，以配合 C# 習慣 */
  private toPascalCase(
    obj: Record<string, unknown>
  ): Record<string, unknown> {
    const result: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(obj)) {
      const pascalKey = key.charAt(0).toUpperCase() + key.slice(1);
      result[pascalKey] = value;
    }
    return result;
  }
}
