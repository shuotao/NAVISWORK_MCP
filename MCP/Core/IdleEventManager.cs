using System;
using System.Collections.Concurrent;
using Autodesk.Navisworks.Api;

namespace NavisworksMCP.Core
{
    /// <summary>
    /// 使用 Navisworks Idle 事件在主執行緒安全執行命令
    /// Navisworks 沒有 Revit 的 ExternalEvent，改用 Application.Idle
    /// </summary>
    public class IdleEventManager
    {
        private static IdleEventManager _instance;
        private static readonly object _lock = new object();

        private readonly ConcurrentQueue<Action> _pendingActions = new ConcurrentQueue<Action>();
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

        /// <summary>
        /// 註冊 Idle 事件監聽
        /// </summary>
        public void Register()
        {
            if (_isRegistered) return;
            Autodesk.Navisworks.Api.Application.Idle += OnIdle;
            _isRegistered = true;
            Logger.Info("Idle 事件已註冊");
        }

        /// <summary>
        /// 取消註冊
        /// </summary>
        public void Unregister()
        {
            if (!_isRegistered) return;
            Autodesk.Navisworks.Api.Application.Idle -= OnIdle;
            _isRegistered = false;
            Logger.Info("Idle 事件已取消註冊");
        }

        /// <summary>
        /// 排入要在主執行緒執行的動作
        /// </summary>
        public void EnqueueAction(Action action)
        {
            _pendingActions.Enqueue(action);
        }

        private void OnIdle(object sender, EventArgs e)
        {
            while (_pendingActions.TryDequeue(out var action))
            {
                try
                {
                    action.Invoke();
                }
                catch (Exception ex)
                {
                    Logger.Error("Idle 執行命令失敗", ex);
                }
            }
        }
    }
}
