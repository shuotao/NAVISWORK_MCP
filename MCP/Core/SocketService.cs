using System;
using System.Diagnostics;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NavisworksMCP.Models;

namespace NavisworksMCP.Core
{
    /// <summary>
    /// WebSocket 服務 — 監聽 port 5150，接收 MCP Server 指令
    /// </summary>
    public class SocketService
    {
        private HttpListener _httpListener;
        private WebSocket _clientSocket;
        private CancellationTokenSource _cts;
        private readonly int _port;

        public bool IsRunning { get; private set; }

        public event EventHandler<NavisCommandRequest> CommandReceived;

        public SocketService(int port = 5150)
        {
            _port = port;
        }

        public async Task StartAsync()
        {
            if (IsRunning) return;

            // 檢查 port 是否被佔用
            if (IsPortInUse(_port))
            {
                Logger.Warn($"Port {_port} 已被佔用，嘗試釋放...");
                TryAutoKillPortOccupant(_port);
                await Task.Delay(1000);
            }

            _cts = new CancellationTokenSource();
            _httpListener = new HttpListener();
            _httpListener.Prefixes.Add($"http://localhost:{_port}/");

            try
            {
                _httpListener.Start();
                IsRunning = true;
                Logger.Info($"NavisworksMCP WebSocket 服務已啟動於 port {_port}");

                _ = AcceptConnectionsAsync(_cts.Token);
            }
            catch (Exception ex)
            {
                Logger.Error("啟動 WebSocket 服務失敗", ex);
                throw;
            }
        }

        private async Task AcceptConnectionsAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested && IsRunning)
            {
                try
                {
                    var context = await _httpListener.GetContextAsync();

                    if (context.Request.IsWebSocketRequest)
                    {
                        var wsContext = await context.AcceptWebSocketAsync(null);
                        _clientSocket = wsContext.WebSocket;
                        Logger.Info("MCP Server 已連接");
                        await ReceiveMessagesAsync(_clientSocket, ct);
                    }
                    else
                    {
                        context.Response.StatusCode = 400;
                        context.Response.Close();
                    }
                }
                catch (ObjectDisposedException) { break; }
                catch (Exception ex)
                {
                    if (!ct.IsCancellationRequested)
                        Logger.Error("接受連線錯誤", ex);
                }
            }
        }

        private async Task ReceiveMessagesAsync(WebSocket socket, CancellationToken ct)
        {
            var buffer = new byte[1024 * 64]; // 64KB buffer

            while (socket.State == WebSocketState.Open && !ct.IsCancellationRequested)
            {
                try
                {
                    var sb = new StringBuilder();
                    WebSocketReceiveResult result;

                    do
                    {
                        result = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), ct);
                        if (result.MessageType == WebSocketMessageType.Close)
                        {
                            Logger.Info("MCP Server 已斷開連接");
                            return;
                        }
                        sb.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));
                    }
                    while (!result.EndOfMessage);

                    var json = sb.ToString();
                    Logger.Info($"收到命令: {json}");

                    var request = Newtonsoft.Json.JsonConvert.DeserializeObject<NavisCommandRequest>(json);
                    if (request != null)
                    {
                        CommandReceived?.Invoke(this, request);
                    }
                }
                catch (WebSocketException) { break; }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Logger.Error("接收訊息錯誤", ex);
                }
            }
        }

        public async Task SendResponseAsync(NavisCommandResponse response)
        {
            if (_clientSocket?.State != WebSocketState.Open) return;

            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(response);
                var bytes = Encoding.UTF8.GetBytes(json);
                await _clientSocket.SendAsync(
                    new ArraySegment<byte>(bytes),
                    WebSocketMessageType.Text,
                    true,
                    CancellationToken.None);
                Logger.Info($"已回傳結果: {(response.Success ? "成功" : "失敗")}");
            }
            catch (Exception ex)
            {
                Logger.Error("傳送回應錯誤", ex);
            }
        }

        public void Stop()
        {
            IsRunning = false;
            try
            {
                _cts?.Cancel();
                _clientSocket?.CloseAsync(WebSocketCloseStatus.NormalClosure, "服務關閉", CancellationToken.None)
                    .Wait(2000);
            }
            catch { }
            try { _httpListener?.Stop(); } catch { }
            try { _httpListener?.Close(); } catch { }

            _clientSocket = null;
            _httpListener = null;
            Logger.Info("WebSocket 服務已停止");
        }

        private bool IsPortInUse(int port)
        {
            try
            {
                var listener = new HttpListener();
                listener.Prefixes.Add($"http://localhost:{port}/");
                listener.Start();
                listener.Stop();
                listener.Close();
                return false;
            }
            catch { return true; }
        }

        private void TryAutoKillPortOccupant(int port)
        {
            try
            {
                var psi = new ProcessStartInfo("netstat", $"-ano")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                var proc = Process.Start(psi);
                var output = proc.StandardOutput.ReadToEnd();
                proc.WaitForExit();

                foreach (var line in output.Split('\n'))
                {
                    if (line.Contains($":{port}") && line.Contains("LISTENING"))
                    {
                        var parts = line.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 4 && int.TryParse(parts[4], out int pid))
                        {
                            var target = Process.GetProcessById(pid);
                            var name = target.ProcessName.ToLower();
                            // 安全白名單：只關閉 node 或自身的舊進程
                            if (name == "node" || name == "navismcp")
                            {
                                target.Kill();
                                Logger.Info($"已終止佔用 port {port} 的進程: {name} (PID: {pid})");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("自動釋放 port 失敗", ex);
            }
        }
    }
}
