using System;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using NavisworksMCP.Core;
using NavisworksMCP.Models;

namespace NavisworksMCP
{
    /// <summary>
    /// Navisworks MCP 主 Plugin — AddInPlugin 提供 Ribbon 按鈕啟動服務
    /// </summary>
    [Plugin("NavisworksMCP.Toggle", "NavisMCP",
        DisplayName = "MCP 服務\n(開/關)",
        ToolTip = "啟動或停止 MCP WebSocket 服務 (port 2233)")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ToggleServicePlugin : AddInPlugin
    {
        private static SocketService _socketService;
        private static bool _idleRegistered;

        public static SocketService SocketService => _socketService;

        public override int Execute(params string[] parameters)
        {
            try
            {
                if (_socketService != null && _socketService.IsRunning)
                {
                    StopService();
                    MessageBox.Show("MCP 服務已停止", "NavisworksMCP",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    StartService();
                    MessageBox.Show(
                        "MCP 服務已啟動\n\n" +
                        "WebSocket 端口: 2233\n" +
                        "等待 MCP Server 連接...",
                        "NavisworksMCP",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("切換 MCP 服務失敗", ex);
                MessageBox.Show($"操作失敗: {ex.Message}", "NavisworksMCP",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }

        private void StartService()
        {
            _socketService = new SocketService(2233);
            _socketService.CommandReceived += OnCommandReceived;
            _socketService.StartAsync().ConfigureAwait(false);

            if (!_idleRegistered)
            {
                IdleEventManager.Instance.Register();
                _idleRegistered = true;
            }

            Logger.Info("MCP 服務已啟動於 port 2233");
        }

        public static void StopService()
        {
            if (_socketService != null)
            {
                _socketService.Stop();
                _socketService = null;
            }
            IdleEventManager.Instance.Unregister();
            _idleRegistered = false;
            Logger.Info("MCP 服務已停止");
        }

        private static void OnCommandReceived(object sender, NavisCommandRequest request)
        {
            // 在 Navisworks 主執行緒中執行命令
            IdleEventManager.Instance.EnqueueAction(() =>
            {
                try
                {
                    var executor = new CommandExecutor();
                    var response = executor.ExecuteCommand(request);
                    _socketService?.SendResponseAsync(response).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    var errorResponse = new NavisCommandResponse
                    {
                        Success = false,
                        Error = ex.Message,
                        RequestId = request.RequestId
                    };
                    _socketService?.SendResponseAsync(errorResponse).ConfigureAwait(false);
                }
            });
        }
    }

    /// <summary>
    /// 開啟日誌檔案
    /// </summary>
    [Plugin("NavisworksMCP.OpenLog", "NavisMCP",
        DisplayName = "開啟\n日誌",
        ToolTip = "開啟 MCP 執行日誌")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class OpenLogPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                var logPath = Logger.GetLogFilePath();
                if (System.IO.File.Exists(logPath))
                    System.Diagnostics.Process.Start("notepad.exe", logPath);
                else
                    MessageBox.Show("日誌檔案尚未建立", "NavisworksMCP");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"開啟日誌失敗: {ex.Message}", "NavisworksMCP");
            }
            return 0;
        }
    }

    /// <summary>
    /// EventWatcher — 監聽文件關閉事件，自動停止服務
    /// </summary>
    [Plugin("NavisworksMCP.EventWatcher", "NavisMCP")]
    [AddInPlugin(AddInLocation.None)]
    public class MCPEventWatcher : EventWatcherPlugin
    {
        public override void OnLoaded()
        {
            Logger.Info("NavisworksMCP EventWatcher 已載入");
        }

        public override void OnUnloading()
        {
            Logger.Info("NavisworksMCP 正在卸載，停止服務...");
            ToggleServicePlugin.StopService();
        }
    }
}
