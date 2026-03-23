namespace NavisworksMCP.Models
{
    /// <summary>
    /// 回傳給 MCP Server 的命令回應
    /// </summary>
    public class NavisCommandResponse
    {
        public bool Success { get; set; }
        public object Data { get; set; }
        public string Error { get; set; }
        public string RequestId { get; set; }
    }
}
