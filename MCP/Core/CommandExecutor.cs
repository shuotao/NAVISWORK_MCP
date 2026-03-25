using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using NavisworksMCP.Models;

namespace NavisworksMCP.Core
{
    /// <summary>
    /// 命令執行器 — 將 MCP 命令轉譯為 Navisworks API 呼叫
    /// </summary>
    public class CommandExecutor
    {
        private readonly Document _doc;

        public CommandExecutor()
        {
            _doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        }

        public NavisCommandResponse ExecuteCommand(NavisCommandRequest request)
        {
            try
            {
                if (_doc == null)
                {
                    return new NavisCommandResponse
                    {
                        Success = false,
                        Error = "沒有打開的文件",
                        RequestId = request.RequestId
                    };
                }

                object result;
                switch (request.Command)
                {
                    case "get_document_info":
                        result = GetDocumentInfo();
                        break;
                    case "get_model_info":
                        result = GetModelInfo();
                        break;
                    case "get_current_selection":
                        result = GetCurrentSelection();
                        break;
                    case "select_items_by_search":
                        result = SelectItemsBySearch(request.Parameters);
                        break;
                    case "get_item_properties":
                        result = GetItemProperties(request.Parameters);
                        break;
                    case "get_model_tree":
                        result = GetModelTree(request.Parameters);
                        break;
                    case "get_viewpoints":
                        result = GetViewpoints();
                        break;
                    case "set_active_viewpoint":
                        result = SetActiveViewpoint(request.Parameters);
                        break;
                    case "get_clash_tests":
                        result = GetClashTests();
                        break;
                    case "get_clash_results":
                        result = GetClashResults(request.Parameters);
                        break;
                    case "get_selection_sets":
                        result = GetSelectionSets();
                        break;
                    case "select_items_by_set":
                        result = SelectItemsBySet(request.Parameters);
                        break;
                    case "clear_selection":
                        result = ClearSelection();
                        break;
                    case "get_all_categories":
                        result = GetAllCategories();
                        break;
                    case "search_items":
                        result = SearchItems(request.Parameters);
                        break;
                    case "get_item_geometry_info":
                        result = GetItemGeometryInfo(request.Parameters);
                        break;
                    case "zoom_to_selection":
                        result = ZoomToSelection();
                        break;
                    case "set_item_override_color":
                        result = SetItemOverrideColor(request.Parameters);
                        break;
                    case "clear_override_colors":
                        result = ClearOverrideColors();
                        break;
                    case "hide_items":
                        result = HideItems(request.Parameters);
                        break;
                    case "unhide_all":
                        result = UnhideAll();
                        break;
                    case "hide_all_except":
                        result = HideAllExcept(request.Parameters);
                        break;
                    case "isolate_selection":
                        result = IsolateSelection();
                        break;
                    case "isolate_by_property":
                        result = IsolateByProperty(request.Parameters);
                        break;
                    case "save_viewpoint":
                        result = SaveViewpoint(request.Parameters);
                        break;
                    // ─── 新增：篩選與資料抽取工具 ───
                    case "get_selection_set_items":
                        result = GetSelectionSetItems(request.Parameters);
                        break;
                    case "execute_search_set":
                        result = ExecuteSearchSet(request.Parameters);
                        break;
                    case "get_override_status":
                        result = GetOverrideStatus(request.Parameters);
                        break;
                    case "get_hidden_items":
                        result = GetHiddenItems(request.Parameters);
                        break;
                    case "get_frozen_items":
                        result = GetFrozenItems(request.Parameters);
                        break;
                    case "batch_get_properties":
                        result = BatchGetProperties(request.Parameters);
                        break;
                    case "get_model_statistics":
                        result = GetModelStatistics(request.Parameters);
                        break;
                    case "scan_subtree":
                        result = ScanSubtree(request.Parameters);
                        break;
                    case "select_subtree":
                        result = SelectSubtree(request.Parameters);
                        break;
                    // ─── Quantification 命令 ───
                    case "quantification_get_items":
                    case "quantification_get_item_groups":
                    case "quantification_import_bq":
                    case "quantification_debug_tables":
                    case "quantification_initialize":
                    case "quantification_exec_sql":
                        return new QuantificationExecutor().Execute(request);
                    default:
                        return new NavisCommandResponse
                        {
                            Success = false,
                            Error = $"未知命令: {request.Command}",
                            RequestId = request.RequestId
                        };
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
                Logger.Error($"執行命令 {request.Command} 失敗", ex);
                return new NavisCommandResponse
                {
                    Success = false,
                    Error = ex.Message,
                    RequestId = request.RequestId
                };
            }
        }

        #region Document & Model Info

        private object GetDocumentInfo()
        {
            return new
            {
                title = _doc.Title,
                fileName = _doc.FileName,
                currentUnits = _doc.Units.ToString(),
                modelCount = _doc.Models.Count,
                isClear = _doc.IsClear
            };
        }

        private object GetModelInfo()
        {
            var models = new List<object>();
            foreach (Model model in _doc.Models)
            {
                var rootItem = model.RootItem;
                int childCount = 0;
                string rootName = null;
                if (rootItem != null)
                {
                    rootName = rootItem.DisplayName;
                    childCount = rootItem.Children.Count();
                }
                models.Add(new
                {
                    fileName = model.FileName,
                    sourceFileName = model.SourceFileName,
                    rootItemDisplayName = rootName,
                    childCount
                });
            }
            return new { models, totalModels = _doc.Models.Count };
        }

        #endregion

        #region Selection

        private object GetCurrentSelection()
        {
            var items = new List<object>();
            foreach (ModelItem item in _doc.CurrentSelection.SelectedItems)
            {
                items.Add(ExtractBasicItemInfo(item));
            }
            return new { selectedCount = items.Count, items };
        }

        private object ClearSelection()
        {
            _doc.CurrentSelection.Clear();
            return new { message = "選擇已清除" };
        }

        private object SelectItemsBySearch(Dictionary<string, object> parameters)
        {
            var categoryName = parameters.ContainsKey("category") ? parameters["category"]?.ToString() : null;
            var propertyName = parameters.ContainsKey("property") ? parameters["property"]?.ToString() : null;
            var value = parameters.ContainsKey("value") ? parameters["value"]?.ToString() : null;

            if (string.IsNullOrEmpty(categoryName) || string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(value))
                throw new ArgumentException("需要 category, property, value 參數");

            var search = new Search();
            search.Selection.SelectAll();
            search.SearchConditions.Add(
                SearchCondition.HasPropertyByDisplayName(categoryName, propertyName)
                    .EqualValue(VariantData.FromDisplayString(value)));

            var found = search.FindAll(_doc, false);
            _doc.CurrentSelection.CopyFrom(found);

            return new { foundCount = found.Count(), message = $"已選擇 {found.Count()} 個項目" };
        }

        private object SelectItemsBySet(Dictionary<string, object> parameters)
        {
            var setName = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(setName))
                throw new ArgumentException("需要 name 參數");

            var sets = _doc.SelectionSets.Value;
            var foundSet = FindSelectionSetByName(sets, setName);
            if (foundSet == null)
                throw new Exception($"找不到選擇集: {setName}");

            if (foundSet is SelectionSet selSet && selSet.HasExplicitModelItems)
            {
                _doc.CurrentSelection.CopyFrom(selSet.ExplicitModelItems);
                return new { message = $"已選擇集合 '{setName}' 的項目" };
            }

            return new { message = "選擇集為空或為搜索集" };
        }

        #endregion

        #region Properties

        private object GetItemProperties(Dictionary<string, object> parameters)
        {
            ModelItem item = GetTargetItem(parameters);
            if (item == null)
                throw new Exception("找不到指定的項目，請先選擇項目或提供搜索條件");

            var categories = new List<object>();
            foreach (PropertyCategory cat in item.PropertyCategories)
            {
                var props = new List<object>();
                foreach (DataProperty prop in cat.Properties)
                {
                    string val = null;
                    try { val = prop.Value?.IsDisplayString == true ? prop.Value.ToDisplayString() : prop.Value?.ToString(); }
                    catch { try { val = prop.Value?.ToString(); } catch { val = "(unreadable)"; } }

                    props.Add(new
                    {
                        name = prop.DisplayName,
                        internalName = prop.Name,
                        value = val,
                        dataType = prop.Value?.DataType.ToString()
                    });
                }
                categories.Add(new
                {
                    categoryName = cat.DisplayName,
                    internalName = cat.Name,
                    properties = props
                });
            }

            return new
            {
                displayName = item.DisplayName,
                classDisplayName = item.ClassDisplayName,
                categories
            };
        }

        private object GetAllCategories()
        {
            var categorySet = new HashSet<string>();

            foreach (ModelItem item in _doc.CurrentSelection.SelectedItems)
            {
                foreach (PropertyCategory cat in item.PropertyCategories)
                {
                    categorySet.Add(cat.DisplayName);
                }
            }

            if (!categorySet.Any())
            {
                // 只掃描前兩層 Children，避免大模型超時
                foreach (Model model in _doc.Models)
                {
                    if (model.RootItem != null)
                    {
                        int scanned = 0;
                        foreach (ModelItem child in model.RootItem.Children)
                        {
                            foreach (PropertyCategory cat in child.PropertyCategories)
                                categorySet.Add(cat.DisplayName);
                            foreach (ModelItem grandchild in child.Children)
                            {
                                foreach (PropertyCategory cat in grandchild.PropertyCategories)
                                    categorySet.Add(cat.DisplayName);
                                if (++scanned >= 30) break;
                            }
                            if (scanned >= 30) break;
                        }
                        break;
                    }
                }
            }

            return new { categories = categorySet.OrderBy(c => c).ToList() };
        }

        #endregion

        #region Model Tree

        private object GetModelTree(Dictionary<string, object> parameters)
        {
            int maxDepth = 3;
            if (parameters != null && parameters.ContainsKey("maxDepth"))
                int.TryParse(parameters["maxDepth"]?.ToString(), out maxDepth);

            var tree = new List<object>();
            foreach (Model model in _doc.Models)
            {
                tree.Add(BuildTreeNode(model.RootItem, 0, maxDepth));
            }
            return new { tree };
        }

        private object BuildTreeNode(ModelItem item, int currentDepth, int maxDepth)
        {
            int childCount = item.Children.Count();

            var node = new Dictionary<string, object>
            {
                ["displayName"] = item.DisplayName,
                ["classDisplayName"] = item.ClassDisplayName,
                ["hasGeometry"] = item.HasGeometry,
                ["isHidden"] = item.IsHidden,
                ["childCount"] = childCount
            };

            if (currentDepth < maxDepth && childCount > 0)
            {
                var children = new List<object>();
                foreach (ModelItem child in item.Children)
                {
                    children.Add(BuildTreeNode(child, currentDepth + 1, maxDepth));
                }
                node["children"] = children;
            }

            return node;
        }

        #endregion

        #region Viewpoints

        private object GetViewpoints()
        {
            var viewpoints = new List<object>();
            var savedViewpoints = _doc.SavedViewpoints.Value;

            foreach (SavedItem savedItem in savedViewpoints)
            {
                CollectViewpoints(savedItem, viewpoints, "");
            }

            return new { viewpoints, count = viewpoints.Count };
        }

        private void CollectViewpoints(SavedItem item, List<object> list, string parentPath)
        {
            string path = string.IsNullOrEmpty(parentPath) ? item.DisplayName : $"{parentPath}/{item.DisplayName}";

            if (item is SavedViewpoint vp)
            {
                var viewpoint = vp.Viewpoint;
                list.Add(new
                {
                    name = item.DisplayName,
                    path,
                    hasViewpoint = viewpoint != null,
                    projection = viewpoint?.Projection.ToString()
                });
            }

            if (item is GroupItem group)
            {
                foreach (SavedItem child in group.Children)
                {
                    CollectViewpoints(child, list, path);
                }
            }
        }

        private object SetActiveViewpoint(Dictionary<string, object> parameters)
        {
            var name = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("需要 name 參數");

            var savedViewpoints = _doc.SavedViewpoints.Value;
            var found = FindViewpointByName(savedViewpoints, name);
            if (found == null)
                throw new Exception($"找不到視點: {name}");

            _doc.SavedViewpoints.CurrentSavedViewpoint = found;
            return new { message = $"已切換到視點: {name}" };
        }

        private SavedViewpoint FindViewpointByName(SavedItemCollection items, string name)
        {
            foreach (SavedItem item in items)
            {
                if (item is SavedViewpoint vp && vp.DisplayName == name)
                    return vp;
                if (item is GroupItem group)
                {
                    var found = FindViewpointByName(group.Children, name);
                    if (found != null) return found;
                }
            }
            return null;
        }

        #endregion

        #region Clash Detection (via COM Interop)

        private object GetClashTests()
        {
            try
            {
                var comDoc = ComApiBridge.State;
                if (comDoc == null)
                    return new { tests = new List<object>(), message = "COM API 不可用" };

                var clashPlugin = comDoc as dynamic;
                // 透過 DocumentParts 存取 Clash
                var parts = _doc.GetType().GetProperty("DocumentParts")?.GetValue(_doc);
                if (parts == null)
                    return new { tests = new List<object>(), message = "Clash Detective 不可用" };

                // 嘗試取得 Clash 資訊
                return new { tests = new List<object>(), message = "請使用 Navisworks GUI 執行碰撞檢測，此功能需要 Manage 版本" };
            }
            catch (Exception ex)
            {
                return new { tests = new List<object>(), message = $"Clash Detective 存取失敗: {ex.Message}" };
            }
        }

        private object GetClashResults(Dictionary<string, object> parameters)
        {
            var testName = parameters.ContainsKey("testName") ? parameters["testName"]?.ToString() : null;
            if (string.IsNullOrEmpty(testName))
                throw new ArgumentException("需要 testName 參數");

            return new { testName, results = new List<object>(), message = "請使用 Navisworks GUI 查看碰撞結果" };
        }

        #endregion

        #region Search

        private object SearchItems(Dictionary<string, object> parameters)
        {
            var category = parameters.ContainsKey("category") ? parameters["category"]?.ToString() : null;
            var property = parameters.ContainsKey("property") ? parameters["property"]?.ToString() : null;
            var value = parameters.ContainsKey("value") ? parameters["value"]?.ToString() : null;
            var condition = parameters.ContainsKey("condition") ? parameters["condition"]?.ToString() : "equals";

            var search = new Search();
            search.Selection.SelectAll();

            if (!string.IsNullOrEmpty(category) && !string.IsNullOrEmpty(property) && !string.IsNullOrEmpty(value))
            {
                var propCondition = SearchCondition.HasPropertyByDisplayName(category, property);
                // Navisworks Search API 主要支援 EqualValue
                search.SearchConditions.Add(propCondition.EqualValue(VariantData.FromDisplayString(value)));
            }
            else if (!string.IsNullOrEmpty(category))
            {
                search.SearchConditions.Add(
                    SearchCondition.HasCategoryByDisplayName(category));
            }

            var found = search.FindAll(_doc, false);
            int maxReturn = 100;
            if (parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            var items = new List<object>();
            int count = 0;
            foreach (ModelItem item in found)
            {
                if (count >= maxReturn) break;
                items.Add(ExtractBasicItemInfo(item));
                count++;
            }

            return new
            {
                totalFound = found.Count(),
                returnedCount = items.Count,
                items
            };
        }

        #endregion

        #region Selection Sets

        private object GetSelectionSets()
        {
            var sets = new List<object>();
            CollectSelectionSets(_doc.SelectionSets.Value, sets, "");
            return new { selectionSets = sets, count = sets.Count };
        }

        private void CollectSelectionSets(SavedItemCollection items, List<object> list, string parentPath)
        {
            foreach (SavedItem item in items)
            {
                string path = string.IsNullOrEmpty(parentPath) ? item.DisplayName : $"{parentPath}/{item.DisplayName}";

                if (item is SelectionSet selSet)
                {
                    list.Add(new
                    {
                        name = item.DisplayName,
                        path,
                        hasExplicitItems = selSet.HasExplicitModelItems,
                        isSearch = selSet.HasSearch,
                        itemCount = selSet.HasExplicitModelItems ? selSet.ExplicitModelItems.Count() : 0
                    });
                }
                else if (item is FolderItem folder)
                {
                    CollectSelectionSets(folder.Children, list, path);
                }
            }
        }

        private SavedItem FindSelectionSetByName(SavedItemCollection items, string name)
        {
            foreach (SavedItem item in items)
            {
                if (item.DisplayName == name) return item;
                if (item is FolderItem folder)
                {
                    var found = FindSelectionSetByName(folder.Children, name);
                    if (found != null) return found;
                }
            }
            return null;
        }

        #endregion

        #region Geometry

        private object GetItemGeometryInfo(Dictionary<string, object> parameters)
        {
            ModelItem item = GetTargetItem(parameters);
            if (item == null)
                throw new Exception("找不到指定項目");

            var bbox = item.BoundingBox();
            return new
            {
                displayName = item.DisplayName,
                hasGeometry = item.HasGeometry,
                boundingBox = bbox != null ? new
                {
                    min = new { x = bbox.Min.X, y = bbox.Min.Y, z = bbox.Min.Z },
                    max = new { x = bbox.Max.X, y = bbox.Max.Y, z = bbox.Max.Z },
                    center = new
                    {
                        x = (bbox.Min.X + bbox.Max.X) / 2,
                        y = (bbox.Min.Y + bbox.Max.Y) / 2,
                        z = (bbox.Min.Z + bbox.Max.Z) / 2
                    }
                } : null
            };
        }

        #endregion

        #region Visual Override

        private object ZoomToSelection()
        {
            if (_doc.CurrentSelection.SelectedItems.Count() == 0)
                throw new Exception("沒有選擇任何項目");

            _doc.ActiveView.FocusOnCurrentSelection();
            return new { message = "已縮放至選擇的項目" };
        }

        private object SetItemOverrideColor(Dictionary<string, object> parameters)
        {
            int r = Convert.ToInt32(parameters["r"]);
            int g = Convert.ToInt32(parameters["g"]);
            int b = Convert.ToInt32(parameters["b"]);

            var color = new Color(r / 255.0, g / 255.0, b / 255.0);

            if (_doc.CurrentSelection.SelectedItems.Count() == 0)
                throw new Exception("請先選擇項目");

            _doc.Models.OverridePermanentColor(
                _doc.CurrentSelection.SelectedItems, color);

            return new
            {
                message = $"已將 {_doc.CurrentSelection.SelectedItems.Count()} 個項目顏色覆蓋為 RGB({r},{g},{b})"
            };
        }

        private object ClearOverrideColors()
        {
            if (_doc.CurrentSelection.SelectedItems.Count() == 0)
            {
                foreach (Model model in _doc.Models)
                {
                    _doc.Models.ResetPermanentMaterials(model.RootItem.Descendants);
                }
                return new { message = "已清除所有項目的顏色覆蓋" };
            }

            _doc.Models.ResetPermanentMaterials(_doc.CurrentSelection.SelectedItems);
            return new { message = $"已重設 {_doc.CurrentSelection.SelectedItems.Count()} 個項目" };
        }

        private object HideItems(Dictionary<string, object> parameters)
        {
            if (_doc.CurrentSelection.SelectedItems.Count() == 0)
                throw new Exception("請先選擇項目");

            _doc.Models.SetHidden(_doc.CurrentSelection.SelectedItems, true);
            return new { message = $"已隱藏 {_doc.CurrentSelection.SelectedItems.Count()} 個項目" };
        }

        private object UnhideAll()
        {
            foreach (Model model in _doc.Models)
            {
                _doc.Models.SetHidden(model.RootItem.Descendants, false);
            }
            return new { message = "已取消隱藏所有項目" };
        }

        /// <summary>
        /// 隱藏所有項目，只顯示指定子樹節點
        /// 高效實現：找到共同父節點，只在同層級切換隱藏/顯示
        /// 參數: name (要顯示的節點名稱), names (多個節點名稱陣列)
        /// </summary>
        private object HideAllExcept(Dictionary<string, object> parameters)
        {
            var showNames = new List<string>();
            if (parameters.ContainsKey("name"))
                showNames.Add(parameters["name"]?.ToString());
            if (parameters.ContainsKey("names"))
            {
                var namesObj = parameters["names"];
                if (namesObj is Newtonsoft.Json.Linq.JArray jArr)
                    showNames.AddRange(jArr.Select(x => x.ToString()));
            }

            if (!showNames.Any())
                throw new ArgumentException("需要 name 或 names 參數");

            // 找到所有要顯示的節點
            var showNodes = new List<ModelItem>();
            foreach (var n in showNames)
            {
                var node = FindNodeByName(n, exact: true) ?? FindNodeByName(n, exact: false);
                if (node != null) showNodes.Add(node);
            }

            if (!showNodes.Any())
                throw new Exception($"找不到任何指定節點");

            try
            {
                // 高效策略：找到目標節點的父節點，隱藏父節點下所有兄弟，只顯示目標
                var parent = showNodes[0].Parent;
                if (parent != null)
                {
                    // 收集同層級的兄弟節點（不遍歷後代）
                    var siblings = new List<ModelItem>();
                    foreach (ModelItem child in parent.Children)
                        siblings.Add(child);

                    _doc.Models.SetHidden(siblings, true);
                    _doc.Models.SetHidden(showNodes, false);

                    // 確保祖先鏈可見
                    var ancestor = parent;
                    while (ancestor != null)
                    {
                        _doc.Models.SetHidden(new List<ModelItem> { ancestor }, false);
                        ancestor = ancestor.Parent;
                    }
                }
                else
                {
                    // 根節點情況
                    var allRoots = new List<ModelItem>();
                    foreach (Model model in _doc.Models)
                    {
                        if (model.RootItem != null)
                            allRoots.Add(model.RootItem);
                    }
                    _doc.Models.SetHidden(allRoots, true);
                    _doc.Models.SetHidden(showNodes, false);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("HideAllExcept failed", ex);
                throw new Exception($"隱藏操作失敗: {ex.Message}");
            }

            return new
            {
                message = $"已隱藏同層級項目，僅顯示 {showNodes.Count} 個節點",
                shownNodes = showNodes.Select(n => n.DisplayName).ToList()
            };
        }

        /// <summary>
        /// 隱藏所有，只顯示當前選擇的項目
        /// </summary>
        private object IsolateSelection()
        {
            var selected = _doc.CurrentSelection.SelectedItems;
            if (selected.Count() == 0)
                throw new Exception("請先選擇項目");

            try
            {
                // 隱藏所有模型的根項目
                foreach (Model model in _doc.Models)
                {
                    if (model.RootItem != null)
                        _doc.Models.SetHidden(model.RootItem.Children, true);
                }

                // 顯示選中的項目
                _doc.Models.SetHidden(selected, false);

                // 確保祖先鏈可見
                var ancestors = new HashSet<ModelItem>();
                foreach (ModelItem item in selected)
                {
                    var ancestor = item.Parent;
                    while (ancestor != null)
                    {
                        if (!ancestors.Add(ancestor)) break; // 已處理過
                        ancestor = ancestor.Parent;
                    }
                }
                if (ancestors.Count > 0)
                    _doc.Models.SetHidden(ancestors.ToList(), false);

                return new { message = $"已隔離顯示 {selected.Count()} 個選中項目", isolatedCount = selected.Count() };
            }
            catch (Exception ex)
            {
                Logger.Error("IsolateSelection failed", ex);
                throw new Exception($"隔離操作失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 搜尋指定屬性值的項目，隱藏其他所有
        /// 參數: category, property, value (搜尋條件)
        /// </summary>
        private object IsolateByProperty(Dictionary<string, object> parameters)
        {
            var categoryName = parameters.ContainsKey("category") ? parameters["category"]?.ToString() : null;
            var propertyName = parameters.ContainsKey("property") ? parameters["property"]?.ToString() : null;
            var value = parameters.ContainsKey("value") ? parameters["value"]?.ToString() : null;

            if (string.IsNullOrEmpty(categoryName) || string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(value))
                throw new ArgumentException("需要 category, property, value 參數");

            try
            {
                var search = new Search();
                search.Selection.SelectAll();
                search.SearchConditions.Add(
                    SearchCondition.HasPropertyByDisplayName(categoryName, propertyName)
                        .EqualValue(VariantData.FromDisplayString(value)));

                var found = search.FindAll(_doc, false);
                int foundCount = found.Count();

                if (foundCount == 0)
                    throw new Exception($"找不到 {categoryName}.{propertyName} = {value} 的項目");

                // 隱藏所有根項目
                foreach (Model model in _doc.Models)
                {
                    if (model.RootItem != null)
                        _doc.Models.SetHidden(model.RootItem.Children, true);
                }

                // 顯示找到的項目
                _doc.Models.SetHidden(found, false);

                // 確保祖先鏈可見
                var ancestors = new HashSet<ModelItem>();
                foreach (ModelItem item in found)
                {
                    var ancestor = item.Parent;
                    while (ancestor != null)
                    {
                        if (!ancestors.Add(ancestor)) break;
                        ancestor = ancestor.Parent;
                    }
                }
                if (ancestors.Count > 0)
                    _doc.Models.SetHidden(ancestors.ToList(), false);

                // 同時選擇這些項目
                _doc.CurrentSelection.CopyFrom(found);

                return new
                {
                    message = $"已隔離 {foundCount} 個 {propertyName}={value} 的項目",
                    isolatedCount = foundCount,
                    property = $"{categoryName}.{propertyName}",
                    value
                };
            }
            catch (Exception ex) when (!(ex is ArgumentException))
            {
                Logger.Error("IsolateByProperty failed", ex);
                throw new Exception($"隔離操作失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 儲存當前視圖為 Saved Viewpoint（安全版，相容 NW 2025）
        /// 參數: name (視點名稱)
        /// </summary>
        private object SaveViewpoint(Dictionary<string, object> parameters)
        {
            var name = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("需要 name 參數");

            try
            {
                var currentViewpoint = _doc.CurrentViewpoint.CreateCopy();
                var savedVp = new SavedViewpoint(currentViewpoint);
                savedVp.DisplayName = name;
                _doc.SavedViewpoints.AddCopy(savedVp);

                return new { message = $"已儲存視點: {name}", viewpointName = name };
            }
            catch (Exception ex)
            {
                Logger.Error("SaveViewpoint failed", ex);
                throw new Exception($"儲存視點失敗: {ex.Message}");
            }
        }

        #endregion

        #region Filter & Data Extraction (新增)

        /// <summary>
        /// 取得 Selection Set 內的所有項目（不選擇），回傳屬性摘要
        /// </summary>
        private object GetSelectionSetItems(Dictionary<string, object> parameters)
        {
            var setName = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(setName))
                throw new ArgumentException("需要 name 參數");

            var sets = _doc.SelectionSets.Value;
            var foundSet = FindSelectionSetByName(sets, setName);
            if (foundSet == null)
                throw new Exception($"找不到選擇集: {setName}");

            int maxReturn = 200;
            if (parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            var items = new List<object>();

            if (foundSet is SelectionSet selSet)
            {
                if (selSet.HasExplicitModelItems)
                {
                    int count = 0;
                    foreach (ModelItem item in selSet.ExplicitModelItems)
                    {
                        if (count >= maxReturn) break;
                        items.Add(ExtractBasicItemInfo(item));
                        count++;
                    }
                    return new { setName, totalItems = selSet.ExplicitModelItems.Count(), returnedCount = items.Count, items };
                }
                else if (selSet.HasSearch)
                {
                    // 這是一個 Search Set，執行搜尋
                    var found = selSet.Search.FindAll(_doc, false);
                    int count = 0;
                    foreach (ModelItem item in found)
                    {
                        if (count >= maxReturn) break;
                        items.Add(ExtractBasicItemInfo(item));
                        count++;
                    }
                    return new { setName, type = "SearchSet", totalItems = found.Count(), returnedCount = items.Count, items };
                }
            }

            return new { setName, totalItems = 0, items, message = "選擇集為空" };
        }

        /// <summary>
        /// 執行已儲存的搜尋集 (Search Set)
        /// </summary>
        private object ExecuteSearchSet(Dictionary<string, object> parameters)
        {
            var setName = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(setName))
                throw new ArgumentException("需要 name 參數");

            bool selectResults = true;
            if (parameters.ContainsKey("select"))
                bool.TryParse(parameters["select"]?.ToString(), out selectResults);

            var sets = _doc.SelectionSets.Value;
            var foundSet = FindSelectionSetByName(sets, setName);
            if (foundSet == null)
                throw new Exception($"找不到選擇集: {setName}");

            if (!(foundSet is SelectionSet selSet) || !selSet.HasSearch)
                throw new Exception($"'{setName}' 不是搜尋集 (Search Set)");

            var found = selSet.Search.FindAll(_doc, false);
            if (selectResults)
                _doc.CurrentSelection.CopyFrom(found);

            int maxReturn = 200;
            if (parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            var items = new List<object>();
            int count = 0;
            foreach (ModelItem item in found)
            {
                if (count >= maxReturn) break;
                items.Add(ExtractBasicItemInfo(item));
                count++;
            }

            return new
            {
                setName,
                totalFound = found.Count(),
                returnedCount = items.Count,
                selected = selectResults,
                items
            };
        }

        /// <summary>
        /// 偵測模型中有顏色覆蓋/透明度覆蓋的項目
        /// 透過比對原始材質與當前狀態來判定
        /// </summary>
        private object GetOverrideStatus(Dictionary<string, object> parameters)
        {
            string scope = "selection"; // selection 或 all
            if (parameters != null && parameters.ContainsKey("scope"))
                scope = parameters["scope"]?.ToString();

            int maxReturn = 200;
            if (parameters != null && parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            IEnumerable<ModelItem> sourceItems;
            if (scope == "all")
            {
                sourceItems = SafeDescendants(5000);
            }
            else
            {
                if (_doc.CurrentSelection.SelectedItems.Count() == 0)
                    throw new Exception("scope=selection 時需先選擇項目，或使用 scope=all");
                sourceItems = _doc.CurrentSelection.SelectedItems;
            }

            var overridden = new List<object>();
            var hidden = new List<object>();
            var frozen = new List<object>();
            int totalScanned = 0;

            foreach (ModelItem item in sourceItems)
            {
                totalScanned++;
                if (overridden.Count >= maxReturn && hidden.Count >= maxReturn) break;

                if (item.IsHidden && hidden.Count < maxReturn)
                {
                    hidden.Add(new { displayName = item.DisplayName, path = GetItemPath(item) });
                }

                if (item.IsFrozen && frozen.Count < maxReturn)
                {
                    frozen.Add(new { displayName = item.DisplayName, path = GetItemPath(item) });
                }

                // 偵測顏色覆蓋：檢查 Geometry 是否有被覆蓋
                // NW API 無直接 "hasOverride" 屬性，但可用 IsRequired 間接判斷
                // 或透過 DocumentModels.IsHidden/IsFrozen 集合方法
            }

            return new
            {
                totalScanned,
                hiddenCount = hidden.Count,
                hiddenItems = hidden,
                frozenCount = frozen.Count,
                frozenItems = frozen,
                message = "使用 set_item_override_color 設定的覆蓋需透過比對原始材質判定，建議用 Selection Sets 管理"
            };
        }

        /// <summary>
        /// 取得模型中所有被隱藏的項目
        /// </summary>
        private object GetHiddenItems(Dictionary<string, object> parameters)
        {
            int maxReturn = 500;
            if (parameters != null && parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            // 優先使用 scope 參數
            string scope = "all";
            if (parameters != null && parameters.ContainsKey("scope"))
                scope = parameters["scope"]?.ToString();

            var hiddenItems = new List<object>();
            IEnumerable<ModelItem> source;

            if (scope == "selection" && _doc.CurrentSelection.SelectedItems.Count() > 0)
            {
                source = _doc.CurrentSelection.SelectedItems;
            }
            else
            {
                source = SafeDescendants(5000);
            }

            int count = 0;
            int totalHidden = 0;
            foreach (ModelItem item in source)
            {
                if (item.IsHidden)
                {
                    totalHidden++;
                    if (count < maxReturn)
                    {
                        hiddenItems.Add(new
                        {
                            displayName = item.DisplayName,
                            classDisplayName = item.ClassDisplayName,
                            path = GetItemPath(item),
                            hasGeometry = item.HasGeometry
                        });
                        count++;
                    }
                }
            }

            return new
            {
                totalHidden,
                returnedCount = hiddenItems.Count,
                items = hiddenItems
            };
        }

        /// <summary>
        /// 取得模型中所有被凍結的項目
        /// </summary>
        private object GetFrozenItems(Dictionary<string, object> parameters)
        {
            int maxReturn = 500;
            if (parameters != null && parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            var frozenItems = new List<object>();
            int totalFrozen = 0;
            int count = 0;

            foreach (ModelItem item in SafeDescendants(5000))
            {
                if (item.IsFrozen)
                {
                    totalFrozen++;
                    if (count < maxReturn)
                    {
                        frozenItems.Add(new
                        {
                            displayName = item.DisplayName,
                            classDisplayName = item.ClassDisplayName,
                            path = GetItemPath(item)
                        });
                        count++;
                    }
                }
            }

            return new
            {
                totalFrozen,
                returnedCount = frozenItems.Count,
                items = frozenItems
            };
        }

        /// <summary>
        /// 批量取得多個項目的屬性 — 用於資料抽取
        /// 可指定要讀取的分類和屬性名稱
        /// </summary>
        private object BatchGetProperties(Dictionary<string, object> parameters)
        {
            // 目標欄位：指定要抽取的 category.property
            var fieldsList = new List<string>();
            if (parameters.ContainsKey("fields"))
            {
                var fieldsObj = parameters["fields"];
                if (fieldsObj is Newtonsoft.Json.Linq.JArray jArr)
                    fieldsList = jArr.Select(x => x.ToString()).ToList();
                else if (fieldsObj is string s)
                    fieldsList = s.Split(',').Select(x => x.Trim()).ToList();
            }

            int maxReturn = 500;
            if (parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out maxReturn);

            // 資料來源：當前選擇 或 指定 Selection Set
            IEnumerable<ModelItem> sourceItems;
            string source = "selection";

            if (parameters.ContainsKey("setName"))
            {
                var setName = parameters["setName"]?.ToString();
                var foundSet = FindSelectionSetByName(_doc.SelectionSets.Value, setName);
                if (foundSet is SelectionSet selSet)
                {
                    if (selSet.HasExplicitModelItems)
                        sourceItems = selSet.ExplicitModelItems;
                    else if (selSet.HasSearch)
                        sourceItems = selSet.Search.FindAll(_doc, false);
                    else
                        throw new Exception($"選擇集 '{setName}' 為空");
                    source = $"set:{setName}";
                }
                else
                    throw new Exception($"找不到選擇集: {setName}");
            }
            else
            {
                if (_doc.CurrentSelection.SelectedItems.Count() == 0)
                    throw new Exception("請先選擇項目或指定 setName 參數");
                sourceItems = _doc.CurrentSelection.SelectedItems;
            }

            var rows = new List<Dictionary<string, object>>();
            int count = 0;
            int totalItems = 0;

            foreach (ModelItem item in sourceItems)
            {
                totalItems++;
                if (count >= maxReturn) continue; // 繼續計數但不加入結果

                var row = new Dictionary<string, object>
                {
                    ["displayName"] = item.DisplayName,
                    ["classDisplayName"] = item.ClassDisplayName,
                    ["path"] = GetItemPath(item)
                };

                if (fieldsList.Any())
                {
                    // 只抽取指定欄位
                    foreach (string field in fieldsList)
                    {
                        var parts = field.Split('.');
                        string catName = parts.Length > 1 ? parts[0] : null;
                        string propName = parts.Length > 1 ? parts[1] : parts[0];

                        string val = null;
                        foreach (PropertyCategory cat in item.PropertyCategories)
                        {
                            if (catName != null && cat.DisplayName != catName)
                                continue;
                            foreach (DataProperty prop in cat.Properties)
                            {
                                if (prop.DisplayName == propName)
                                {
                                    try { val = prop.Value?.IsDisplayString == true ? prop.Value.ToDisplayString() : prop.Value?.ToString(); }
                                    catch { try { val = prop.Value?.ToString(); } catch { } }
                                    break;
                                }
                            }
                            if (val != null) break;
                        }
                        row[field] = val;
                    }
                }
                else
                {
                    // 抽取所有屬性（扁平化）
                    foreach (PropertyCategory cat in item.PropertyCategories)
                    {
                        foreach (DataProperty prop in cat.Properties)
                        {
                            var key = $"{cat.DisplayName}.{prop.DisplayName}";
                            string pv = null;
                            try { pv = prop.Value?.IsDisplayString == true ? prop.Value.ToDisplayString() : prop.Value?.ToString(); }
                            catch { try { pv = prop.Value?.ToString(); } catch { } }
                            row[key] = pv;
                        }
                    }
                }

                rows.Add(row);
                count++;
            }

            return new
            {
                source,
                totalItems,
                returnedCount = rows.Count,
                fields = fieldsList.Any() ? fieldsList : null,
                rows
            };
        }

        /// <summary>
        /// 模型統計 — 按分類/圖層/來源檔案彙總數量
        /// </summary>
        private object GetModelStatistics(Dictionary<string, object> parameters)
        {
            string groupBy = "classDisplayName";
            if (parameters != null && parameters.ContainsKey("groupBy"))
                groupBy = parameters["groupBy"]?.ToString();

            bool geometryOnly = true;
            if (parameters != null && parameters.ContainsKey("geometryOnly"))
                bool.TryParse(parameters["geometryOnly"]?.ToString(), out geometryOnly);

            var stats = new Dictionary<string, int>();
            int totalItems = 0;
            int totalWithGeometry = 0;
            int totalHidden = 0;

            foreach (ModelItem item in SafeDescendants(10000))
            {
                totalItems++;
                if (item.HasGeometry) totalWithGeometry++;
                if (item.IsHidden) totalHidden++;

                if (geometryOnly && !item.HasGeometry) continue;

                string key;
                switch (groupBy)
                {
                    case "className":
                        key = item.ClassName ?? "(無)";
                        break;
                    case "layer":
                        key = item.IsLayer ? item.DisplayName : "(非圖層)";
                        // 取得圖層名稱：往上找到 IsLayer=true 的祖先
                        var layerAncestor = item.Ancestors.FirstOrDefault(a => a.IsLayer);
                        if (layerAncestor != null) key = layerAncestor.DisplayName;
                        break;
                    case "sourceFile":
                        key = item.Model?.SourceFileName ?? item.Model?.FileName ?? "(未知)";
                        break;
                    case "classDisplayName":
                    default:
                        key = item.ClassDisplayName ?? "(無)";
                        break;
                }

                if (stats.ContainsKey(key))
                    stats[key]++;
                else
                    stats[key] = 1;
            }

            var sortedStats = stats.OrderByDescending(kv => kv.Value)
                .Select(kv => new { name = kv.Key, count = kv.Value })
                .ToList();

            return new
            {
                totalItems,
                totalWithGeometry,
                totalHidden,
                groupBy,
                categoryCount = sortedStats.Count,
                statistics = sortedStats
            };
        }

        #endregion

        #region Helpers

        private ModelItem GetTargetItem(Dictionary<string, object> parameters)
        {
            // 先看當前選擇
            if (_doc.CurrentSelection.SelectedItems.Count() > 0)
                return _doc.CurrentSelection.SelectedItems.First;

            // 如果有搜索參數
            if (parameters != null && parameters.ContainsKey("category") && parameters.ContainsKey("property"))
            {
                var search = new Search();
                search.Selection.SelectAll();
                search.SearchConditions.Add(
                    SearchCondition.HasPropertyByDisplayName(
                        parameters["category"].ToString(),
                        parameters["property"].ToString())
                    .EqualValue(VariantData.FromDisplayString(parameters["value"]?.ToString())));
                return search.FindFirst(_doc, false);
            }

            return null;
        }

        private object ExtractBasicItemInfo(ModelItem item)
        {
            return new
            {
                displayName = item.DisplayName,
                classDisplayName = item.ClassDisplayName,
                hasGeometry = item.HasGeometry,
                isHidden = item.IsHidden,
                instanceGuid = item.InstanceGuid.ToString(),
                ancestorPath = GetItemPath(item)
            };
        }

        /// <summary>
        /// 安全遍歷模型項目，限制最大數量避免大模型卡死
        /// </summary>
        private IEnumerable<ModelItem> SafeDescendants(int maxItems)
        {
            int count = 0;
            foreach (Model model in _doc.Models)
            {
                if (model.RootItem == null) continue;
                foreach (ModelItem item in model.RootItem.Descendants)
                {
                    if (count >= maxItems) yield break;
                    count++;
                    yield return item;
                }
            }
        }

        /// <summary>
        /// 從 ModelItem 讀取 Element.Category/Family/Type，找到就寫入 ref 參數
        /// </summary>
        private void ReadElementProps(ModelItem item, ref string category, ref string family, ref string type)
        {
            foreach (PropertyCategory cat in item.PropertyCategories)
            {
                if (cat.DisplayName != "Element") continue;
                foreach (DataProperty prop in cat.Properties)
                {
                    try
                    {
                        var val = prop.Value?.IsDisplayString == true
                            ? prop.Value.ToDisplayString()
                            : prop.Value?.ToString();
                        if (string.IsNullOrEmpty(val)) continue;
                        switch (prop.DisplayName)
                        {
                            case "Category": if (string.IsNullOrEmpty(category)) category = val; break;
                            case "Family": if (string.IsNullOrEmpty(family)) family = val; break;
                            case "Type": if (string.IsNullOrEmpty(type)) type = val; break;
                        }
                    }
                    catch { }
                }
                break;
            }
        }

        private string GetItemPath(ModelItem item)
        {
            if (item == null) return "";
            var parts = new List<string>();
            var current = item;
            while (current != null)
            {
                if (!string.IsNullOrEmpty(current.DisplayName))
                    parts.Insert(0, current.DisplayName);
                current = current.Parent;
            }
            return string.Join(" > ", parts);
        }

        /// <summary>
        /// 在模型樹中找到指定名稱的節點
        /// 支援部分匹配和精確匹配
        /// </summary>
        private ModelItem FindNodeByName(string name, bool exact = false)
        {
            foreach (Model model in _doc.Models)
            {
                if (model.RootItem == null) continue;
                var found = FindNodeRecursive(model.RootItem, name, exact, 0, 5);
                if (found != null) return found;
            }
            return null;
        }

        private ModelItem FindNodeRecursive(ModelItem parent, string name, bool exact, int depth, int maxDepth)
        {
            if (depth > maxDepth) return null;
            foreach (ModelItem child in parent.Children)
            {
                var dn = child.DisplayName ?? "";
                if (exact ? dn == name : dn.Contains(name))
                    return child;
                var found = FindNodeRecursive(child, name, exact, depth + 1, maxDepth);
                if (found != null) return found;
            }
            return null;
        }

        /// <summary>
        /// 掃描指定子樹節點下的所有幾何元素，按 Revit Category 和 Family/Type 分組統計
        /// 參數: name (節點名稱), maxItems (最大掃描數，預設 50000), fields (額外屬性欄位)
        /// </summary>
        private object ScanSubtree(Dictionary<string, object> parameters)
        {
            var name = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("需要 name 參數（NWC/NWD 節點名稱）");

            int maxItems = 50000;
            if (parameters.ContainsKey("maxItems"))
                int.TryParse(parameters["maxItems"]?.ToString(), out maxItems);

            // 額外要讀取的屬性欄位
            var fieldsList = new List<string>();
            if (parameters.ContainsKey("fields"))
            {
                var fieldsObj = parameters["fields"];
                if (fieldsObj is Newtonsoft.Json.Linq.JArray jArr)
                    fieldsList = jArr.Select(x => x.ToString()).ToList();
                else if (fieldsObj is string s)
                    fieldsList = s.Split(',').Select(x => x.Trim()).ToList();
            }

            var node = FindNodeByName(name, exact: true) ?? FindNodeByName(name, exact: false);
            if (node == null)
                throw new Exception($"找不到節點: {name}");

            // 統計資料結構
            var categoryStats = new Dictionary<string, Dictionary<string, int>>(); // category → { family|type → count }
            int totalGeometry = 0;
            int totalItems = 0;
            var rows = new List<Dictionary<string, object>>();

            bool collectRows = fieldsList.Any();
            int rowLimit = 5000;
            if (parameters.ContainsKey("maxResults"))
                int.TryParse(parameters["maxResults"]?.ToString(), out rowLimit);

            foreach (ModelItem item in node.Descendants)
            {
                totalItems++;
                if (totalItems > maxItems) break;

                if (!item.HasGeometry) continue;
                totalGeometry++;

                // 讀取 Element.Category, Family, Type — 從自身或祖先節點
                string elemCategory = null, elemFamily = null, elemType = null;
                ReadElementProps(item, ref elemCategory, ref elemFamily, ref elemType);

                // 幾何節點通常沒有 Element 屬性，往上找祖先（最多 5 層）
                if (string.IsNullOrEmpty(elemCategory))
                {
                    var ancestor = item.Parent;
                    int depth = 0;
                    while (ancestor != null && depth < 5)
                    {
                        ReadElementProps(ancestor, ref elemCategory, ref elemFamily, ref elemType);
                        if (!string.IsNullOrEmpty(elemCategory)) break;
                        ancestor = ancestor.Parent;
                        depth++;
                    }
                }

                if (string.IsNullOrEmpty(elemCategory)) elemCategory = item.ClassDisplayName ?? "(unknown)";

                // 統計
                if (!categoryStats.ContainsKey(elemCategory))
                    categoryStats[elemCategory] = new Dictionary<string, int>();
                var typeKey = (elemFamily ?? elemCategory) + " | " + (elemType ?? "");
                if (!categoryStats[elemCategory].ContainsKey(typeKey))
                    categoryStats[elemCategory][typeKey] = 0;
                categoryStats[elemCategory][typeKey]++;

                // 收集詳細行（如有指定 fields）
                if (collectRows && rows.Count < rowLimit)
                {
                    var row = new Dictionary<string, object>
                    {
                        ["displayName"] = item.DisplayName,
                        ["path"] = GetItemPath(item),
                        ["Element.Category"] = elemCategory,
                        ["Element.Family"] = elemFamily,
                        ["Element.Type"] = elemType
                    };

                    foreach (string field in fieldsList)
                    {
                        if (row.ContainsKey(field)) continue;
                        var fp = field.Split('.');
                        string catName = fp.Length > 1 ? fp[0] : null;
                        string propName = fp.Length > 1 ? fp[1] : fp[0];
                        string val = null;
                        foreach (PropertyCategory cat in item.PropertyCategories)
                        {
                            if (catName != null && cat.DisplayName != catName) continue;
                            foreach (DataProperty prop in cat.Properties)
                            {
                                if (prop.DisplayName == propName)
                                {
                                    try { val = prop.Value?.IsDisplayString == true ? prop.Value.ToDisplayString() : prop.Value?.ToString(); }
                                    catch { }
                                    break;
                                }
                            }
                            if (val != null) break;
                        }
                        row[field] = val;
                    }
                    rows.Add(row);
                }
            }

            // 組裝結果
            var stats = categoryStats.Select(kvp => new
            {
                category = kvp.Key,
                totalCount = kvp.Value.Values.Sum(),
                types = kvp.Value.Select(t => new { typeKey = t.Key, count = t.Value })
                    .OrderByDescending(t => t.count).ToList()
            }).OrderByDescending(c => c.totalCount).ToList();

            var result = new Dictionary<string, object>
            {
                ["nodeName"] = node.DisplayName,
                ["nodePath"] = GetItemPath(node),
                ["totalDescendants"] = totalItems,
                ["totalGeometry"] = totalGeometry,
                ["categoryCount"] = stats.Count,
                ["categories"] = stats
            };
            if (collectRows)
            {
                result["rows"] = rows;
                result["rowCount"] = rows.Count;
            }
            return result;
        }

        /// <summary>
        /// 選取指定子樹節點下的所有幾何元素到 CurrentSelection
        /// 參數: name (節點名稱)
        /// </summary>
        private object SelectSubtree(Dictionary<string, object> parameters)
        {
            var name = parameters.ContainsKey("name") ? parameters["name"]?.ToString() : null;
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("需要 name 參數");

            var node = FindNodeByName(name, exact: true) ?? FindNodeByName(name, exact: false);
            if (node == null)
                throw new Exception($"找不到節點: {name}");

            // 收集所有有幾何的後代
            var geometryItems = new List<ModelItem>();
            int total = 0;
            foreach (ModelItem item in node.Descendants)
            {
                total++;
                if (total > 100000) break; // 安全上限
                if (item.HasGeometry)
                    geometryItems.Add(item);
            }

            _doc.CurrentSelection.Clear();
            _doc.CurrentSelection.AddRange(geometryItems);

            return new
            {
                nodeName = node.DisplayName,
                totalDescendants = total,
                selectedCount = geometryItems.Count,
                message = $"已選擇 {geometryItems.Count} 個幾何項目（來自 {node.DisplayName}）"
            };
        }

        #endregion
    }

    /// <summary>
    /// COM API 橋接 — 用於存取 .NET API 不提供的功能
    /// </summary>
    internal static class ComApiBridge
    {
        public static Autodesk.Navisworks.Api.Interop.ComApi.InwOpState10 State
        {
            get
            {
                try
                {
                    return Autodesk.Navisworks.Api.ComApi.ComApiBridge.State;
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
