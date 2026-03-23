using System;
using System.Collections.Concurrent;
using System.Windows.Forms;

namespace NavisworksMCP.Core
{
    /// <summary>
    /// 使用 Timer 輪詢方式在主執行緒安全執行命令
    /// Application.Idle 在大模型中不可靠，改用 Timer
    /// </summary>
    public class IdleEventManager
    {
        private static IdleEventManager _instance;
        private static readonly object _lock = new object();

        private readonly ConcurrentQueue<Action> _pendingActions = new ConcurrentQueue<Action>();
        private Timer _timer;
        private bool _isRegistered;

        private IdleEventManager() { }

        public static IdleEventManager Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                        _instance = new IdleEventManager();
                    return _instance;
                }
            }
        }

        public void Register()
        {
            if (_isRegistered) return;

            _timer = new Timer();
            _timer.Interval = 500; // 每 500ms 檢查一次
            _timer.Tick += OnTimerTick;
            _timer.Start();

            _isRegistered = true;
            Logger.Info("Timer 輪詢已啟動 (500ms)");
        }

        public void Unregister()
        {
            if (!_isRegistered) return;

            if (_timer != null)
            {
                _timer.Stop();
                _timer.Tick -= OnTimerTick;
                _timer.Dispose();
                _timer = null;
            }

            _isRegistered = false;
            Logger.Info("Timer 輪詢已停止");
        }

        public void EnqueueAction(Action action)
        {
            _pendingActions.Enqueue(action);
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            // 每次 tick 只處理一個命令，避免阻塞 UI 太久
            if (_pendingActions.TryDequeue(out var action))
            {
                try
                {
                    _timer.Stop(); // 暫停計時器，避免重入
                    action.Invoke();
                }
                catch (Exception ex)
                {
                    Logger.Error("Timer 執行命令失敗", ex);
                }
                finally
                {
                    _timer.Start(); // 恢復計時器
                }
            }
        }
    }
}
