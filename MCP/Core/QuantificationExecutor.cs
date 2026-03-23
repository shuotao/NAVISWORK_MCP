using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Data;
using Autodesk.Navisworks.Api.Takeoff;
using NavisworksMCP.Models;
using Newtonsoft.Json.Linq;

namespace NavisworksMCP.Core
{
    /// <summary>
    /// Quantification (Takeoff) 命令執行器
    /// 負責操作 Navisworks Quantification 面板
    /// </summary>
    public class QuantificationExecutor
    {
        private readonly Document _doc;

        public QuantificationExecutor()
        {
            _doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        }

        public NavisCommandResponse Execute(NavisCommandRequest request)
        {
            try
            {
                if (_doc == null)
                    return Error("沒有打開的文件", request.RequestId);

                object result;
                switch (request.Command)
                {
                    case "quantification_get_items":
                        result = GetItems();
                        break;
                    case "quantification_get_item_groups":
                        result = GetItemGroups();
                        break;
                    case "quantification_import_bq":
                        result = ImportBQItems(request.Parameters);
                        break;
                    case "quantification_debug_tables":
                        result = DebugListTables();
                        break;
                    case "quantification_initialize":
                        result = InitializeQuantification();
                        break;
                    case "quantification_exec_sql":
                        result = ExecSql(request.Parameters);
                        break;
                    default:
                        return Error($"未知 Quantification 命令: {request.Command}", request.RequestId);
                }

                return new NavisCommandResponse
                {
                    Success = true,
                    Data = result,
                    RequestId = request.RequestId
                };
            }
            catch (Exception ex)
            {
                Logger.Error($"Quantification 命令 {request.Command} 失敗", ex);
                return Error(ex.Message, request.RequestId);
            }
        }

        /// <summary>
        /// 讀取現有的 Quantification Items
        /// </summary>
        private object GetItems()
        {
            var takeoff = GetTakeoff();
            if (takeoff == null)
                return new { items = new List<object>(), message = "Quantification 未初始化，請先開啟 Quantification 面板" };

            var items = new List<object>();
            var itemTable = takeoff.Items;
            var variables = itemTable.Variables;

            // 讀取所有 row
            var db = takeoff.Database;
            var conn = db.Value;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM TK_Item";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var row = new Dictionary<string, object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        }
                        items.Add(row);
                    }
                }
            }

            return new { items, count = items.Count };
        }

        /// <summary>
        /// 讀取 Item Groups
        /// </summary>
        private object GetItemGroups()
        {
            var takeoff = GetTakeoff();
            if (takeoff == null)
                return new { groups = new List<object>(), message = "Quantification 未初始化" };

            var groups = new List<object>();
            var db = takeoff.Database;
            var conn = db.Value;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM TK_ItemGroup";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var row = new Dictionary<string, object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        }
                        groups.Add(row);
                    }
                }
            }

            return new { groups, count = groups.Count };
        }

        /// <summary>
        /// 從 BQ 資料批量匯入 Quantification Items
        /// </summary>
        private object ImportBQItems(Dictionary<string, object> parameters)
        {
            var takeoff = GetTakeoff();
            if (takeoff == null)
                throw new Exception("Quantification 未初始化，請先在 Navisworks 中開啟 Quantification 面板 (Home > Quantification)");

            // 解析 BQ 項目
            var bqItems = new List<BQItem>();
            if (parameters.ContainsKey("items"))
            {
                var itemsObj = parameters["items"];
                JArray jArr;
                if (itemsObj is JArray ja)
                    jArr = ja;
                else
                    jArr = JArray.Parse(itemsObj.ToString());

                foreach (JObject item in jArr)
                {
                    bqItems.Add(new BQItem
                    {
                        Sheet = item["sheet"]?.ToString() ?? "",
                        SNo = item["sno"]?.ToString() ?? "",
                        Description = item["desc"]?.ToString() ?? "",
                        Unit = item["unit"]?.ToString() ?? "",
                        Qty = item["qty"]?.ToObject<double>() ?? 0
                    });
                }
            }

            if (!bqItems.Any())
                throw new ArgumentException("沒有提供 BQ 項目資料");

            var db = takeoff.Database;
            var conn = db.Value;
            int groupsCreated = 0;
            int itemsCreated = 0;
            var errors = new List<string>();

            // 按 sheet 分組建立 ItemGroup，每個 sheet 對應一個群組
            var sheets = bqItems.Select(b => b.Sheet).Distinct().ToList();

            using (var transaction = db.BeginTransaction(DatabaseChangedAction.Reset))
            {
                try
                {
                    // 取得目前最大 CatalogId 避免 UNIQUE 衝突
                    long nextCatalogId = 1;
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COALESCE(MAX(CatalogId), 0) + 1 FROM TK_ItemGroup";
                        nextCatalogId = (long)cmd.ExecuteScalar();
                    }

                    foreach (var sheet in sheets)
                    {
                        // 建立 Item Group (分類)
                        long groupId;
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = @"INSERT INTO TK_ItemGroup (Name, Description, WBS, Parent, Status, CatalogId)
                                              VALUES (@name, @desc, @wbs, 0, 0, @catId)";
                            cmd.Parameters.AddWithValue("@name", sheet);
                            cmd.Parameters.AddWithValue("@desc", $"BQ - {sheet}");
                            cmd.Parameters.AddWithValue("@wbs", sheet.Split('.')[0].Trim());
                            cmd.Parameters.AddWithValue("@catId", nextCatalogId++);
                            cmd.ExecuteNonQuery();
                            groupsCreated++;
                        }

                        // 取得剛建立的 group ID
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT last_insert_rowid()";
                            groupId = (long)cmd.ExecuteScalar();
                        }

                        // 取得 Item 的 CatalogId 起始值
                        long nextItemCatId = 1;
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT COALESCE(MAX(CatalogId), 0) + 1 FROM TK_Item";
                            nextItemCatId = (long)cmd.ExecuteScalar();
                        }

                        // 建立該 sheet 下的 Items
                        var sheetItems = bqItems.Where(b => b.Sheet == sheet).ToList();
                        foreach (var bq in sheetItems)
                        {
                            try
                            {
                                using (var cmd = conn.CreateCommand())
                                {
                                    var itemName = $"{bq.SNo} - {bq.Description}";
                                    if (itemName.Length > 200) itemName = itemName.Substring(0, 200);

                                    cmd.CommandText = @"INSERT INTO TK_Item
                                        (Name, Description, WBS, Parent, Status, CatalogId, Color, Transparency, LineThickness, PrimaryQuantity_Value, PrimaryQuantity_Units)
                                        VALUES (@name, @desc, @wbs, @parent, 0, @catId, 0, 0.0, 1, @qty, @unit)";
                                    cmd.Parameters.AddWithValue("@name", itemName);
                                    cmd.Parameters.AddWithValue("@desc", bq.Description);
                                    cmd.Parameters.AddWithValue("@wbs", bq.SNo);
                                    cmd.Parameters.AddWithValue("@parent", groupId);
                                    cmd.Parameters.AddWithValue("@catId", nextItemCatId++);
                                    cmd.Parameters.AddWithValue("@qty", bq.Qty);
                                    cmd.Parameters.AddWithValue("@unit", bq.Unit);
                                    cmd.ExecuteNonQuery();
                                    itemsCreated++;
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"{bq.SNo}: {ex.Message}");
                            }
                        }
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    Logger.Error("匯入 BQ 失敗，回滾交易", ex);
                    throw;
                }
            }

            return new
            {
                message = $"匯入完成：{groupsCreated} 個群組，{itemsCreated} 個項目",
                groupsCreated,
                itemsCreated,
                totalSheets = sheets.Count,
                errors = errors.Any() ? errors : null
            };
        }

        /// <summary>
        /// 初始化 Quantification — 建立 schema 和表結構
        /// </summary>
        private object InitializeQuantification()
        {
            // 確保 Takeoff 實例存在
            var takeoff = DocumentTakeoff.CreateInstance(_doc);
            if (takeoff == null)
                throw new Exception("無法建立 Takeoff 實例");

            var configMgr = Autodesk.Navisworks.Api.Interop.LcTkConfigManager.Instance;
            if (configMgr == null)
                throw new Exception("無法取得 ConfigManager");

            if (configMgr.IsDatabaseSetUp())
                return new { message = "Quantification 已經初始化", alreadySetup = true };

            // 載入預設範本並建立 schema
            string errorMsg = null;
            bool loaded = configMgr.LoadTemplateConfigFile(out errorMsg);
            if (!loaded)
            {
                Logger.Warn($"LoadTemplateConfigFile: {errorMsg}, 嘗試直接 CreateSchema...");
            }

            bool created = configMgr.CreateSchema(out errorMsg);
            if (!created)
                throw new Exception($"CreateSchema 失敗: {errorMsg}");

            return new { message = "Quantification 初始化成功", created = true };
        }

        private object DebugListTables()
        {
            var takeoff = GetTakeoff();
            if (takeoff == null)
                return new { message = "Quantification 未初始化", initialized = false };

            var tables = new List<object>();
            var db = takeoff.Database;
            var conn = db.Value;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT name, type FROM sqlite_master WHERE type='table' ORDER BY name";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var tableName = reader.GetString(0);
                        // 也取得欄位資訊
                        var columns = new List<string>();
                        using (var cmd2 = conn.CreateCommand())
                        {
                            cmd2.CommandText = $"PRAGMA table_info({tableName})";
                            using (var r2 = cmd2.ExecuteReader())
                            {
                                while (r2.Read())
                                    columns.Add(r2.GetString(1));
                            }
                        }
                        tables.Add(new { name = tableName, columns });
                    }
                }
            }

            return new { initialized = true, tables, tableCount = tables.Count };
        }

        /// <summary>
        /// 通用 SQL 執行 — 以後不需要重啟 NW 就能操作 DB
        /// </summary>
        private object ExecSql(Dictionary<string, object> parameters)
        {
            var takeoff = GetTakeoff();
            if (takeoff == null)
                throw new Exception("Quantification 未初始化");

            var sql = parameters.ContainsKey("sql") ? parameters["sql"]?.ToString() : null;
            if (string.IsNullOrEmpty(sql))
                throw new ArgumentException("需要 sql 參數");

            bool isQuery = sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
                        || sql.TrimStart().StartsWith("PRAGMA", StringComparison.OrdinalIgnoreCase);

            var db = takeoff.Database;
            var conn = db.Value;

            if (isQuery)
            {
                var rows = new List<Dictionary<string, object>>();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var row = new Dictionary<string, object>();
                            for (int i = 0; i < reader.FieldCount; i++)
                                row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            rows.Add(row);
                        }
                    }
                }
                return new { type = "query", rowCount = rows.Count, rows };
            }
            else
            {
                using (var transaction = db.BeginTransaction(DatabaseChangedAction.Reset))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = sql;
                        int affected = cmd.ExecuteNonQuery();
                        transaction.Commit();
                        return new { type = "execute", affectedRows = affected };
                    }
                }
            }
        }

        private DocumentTakeoff GetTakeoff()
        {
            try
            {
                return _doc.GetTakeoff();
            }
            catch
            {
                return null;
            }
        }

        private NavisCommandResponse Error(string msg, string requestId)
        {
            return new NavisCommandResponse { Success = false, Error = msg, RequestId = requestId };
        }

        private class BQItem
        {
            public string Sheet { get; set; }
            public string SNo { get; set; }
            public string Description { get; set; }
            public string Unit { get; set; }
            public double Qty { get; set; }
        }
    }
}
