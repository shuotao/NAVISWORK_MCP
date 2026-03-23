using System.Collections.Generic;

namespace NavisworksMCP.Models
{
    /// <summary>
    /// MCP Server 傳來的命令請求
    /// </summary>
    public class NavisCommandRequest
    {
        public string Command { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public string RequestId { get; set; }
    }
}
