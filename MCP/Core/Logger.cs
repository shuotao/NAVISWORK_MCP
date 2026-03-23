using System;
using System.IO;

namespace NavisworksMCP.Core
{
    public static class Logger
    {
        private static readonly string LogDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "NavisworksMCP", "Logs");

        private static readonly object _lock = new object();

        static Logger()
        {
            if (!Directory.Exists(LogDir))
                Directory.CreateDirectory(LogDir);
        }

        private static string LogFile => Path.Combine(LogDir, $"NavisMCP_{DateTime.Now:yyyy-MM-dd}.log");

        public static void Info(string message) => Write("INFO", message);
        public static void Warn(string message) => Write("WARN", message);
        public static void Error(string message, Exception ex = null)
        {
            var msg = ex != null ? $"{message}: {ex.Message}\n{ex.StackTrace}" : message;
            Write("ERROR", msg);
        }

        private static void Write(string level, string message)
        {
            lock (_lock)
            {
                try
                {
                    File.AppendAllText(LogFile, $"[{DateTime.Now:HH:mm:ss}] [{level}] {message}\n");
                }
                catch { }
            }
        }

        public static string GetLogFilePath() => LogFile;
    }
}
