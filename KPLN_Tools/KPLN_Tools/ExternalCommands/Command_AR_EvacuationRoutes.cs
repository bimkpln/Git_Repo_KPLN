using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.ExternalCommands.UI;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Interop;
using System.Reflection;
using KPLN_Tools.Common;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_EvacuationRoutes : IExternalCommand
    {
        private static UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;

        private enum EvacuationRoutesRequestKind
        {
            None,
            SelectElement,
            Build,
            PickAndBuild
        }

        private sealed class EvacuationRoutesExternalEventHandler : IExternalEventHandler
        {
            private readonly Command_AR_EvacuationRoutes _owner;
            private EvacuationRoutesRequestKind _requestKind = EvacuationRoutesRequestKind.None;
            private EvacuationRoutesDialog _dialog;
            private EvacuationRoutesDialogResult _data;
            private long _elementId;

            public EvacuationRoutesExternalEventHandler(Command_AR_EvacuationRoutes owner)
            {
                _owner = owner;
            }

            public string GetName()
            {
                return "KPLN Evacuation Routes";
            }

            public void RequestSelect(EvacuationRoutesDialog dialog, long elementId)
            {
                _dialog = dialog;
                _elementId = elementId;
                _data = null;
                _requestKind = EvacuationRoutesRequestKind.SelectElement;
            }

            public void RequestBuild(EvacuationRoutesDialog dialog, EvacuationRoutesDialogResult data)
            {
                _dialog = dialog;
                _data = data;
                _elementId = 0;
                _requestKind = EvacuationRoutesRequestKind.Build;
            }

            public void RequestPickAndBuild(EvacuationRoutesDialog dialog, EvacuationRoutesDialogResult data)
            {
                _dialog = dialog;
                _data = data;
                _elementId = 0;
                _requestKind = EvacuationRoutesRequestKind.PickAndBuild;
            }

            public void Execute(UIApplication app)
            {
                EvacuationRoutesRequestKind requestKind = _requestKind;
                EvacuationRoutesDialog dialog = _dialog;
                EvacuationRoutesDialogResult data = _data;
                long elementId = _elementId;

                _requestKind = EvacuationRoutesRequestKind.None;
                _dialog = null;
                _data = null;
                _elementId = 0;

                if (_owner == null || app == null)
                    return;

                uiapp = app;
                _owner.uidoc = app.ActiveUIDocument;
                _owner.doc = _owner.uidoc?.Document;

                try
                {
                    if (requestKind == EvacuationRoutesRequestKind.SelectElement)
                    {
                        SelectAndShowElement(_owner.uidoc, elementId);
                        Notify(dialog, $"Выбран элемент ID {elementId}.");
                        return;
                    }

                    if (requestKind == EvacuationRoutesRequestKind.Build)
                    {
                        EvacuationRoutesOperationResult result = _owner.RunEvacuationRoutesOperation(data);
                        ApplyResult(dialog, result);
                        return;
                    }

                    if (requestKind == EvacuationRoutesRequestKind.PickAndBuild)
                    {
                        long? pickedId;
                        HideForPick(dialog);
                        try
                        {
                            pickedId = PickStairElementId(app, _owner.doc);
                        }
                        finally
                        {
                            RestoreAfterPick(dialog);
                        }

                        if (!pickedId.HasValue)
                        {
                            Finish(dialog, "Окно закрыто.");
                            return;
                        }

                        SelectRow(dialog, pickedId.Value);

                        EvacuationRoutesDialogResult pickedData = new EvacuationRoutesDialogResult(
                            data.HeightMm,
                            data.WidthMm,
                            data.UseRunWidth,
                            true,
                            data.AddToEvacuationWorkset,
                            data.EvacuationWorksetId,
                            pickedId.Value);

                        Notify(dialog, "Идёт обработка маршрутов...");
                        EvacuationRoutesOperationResult result = _owner.RunEvacuationRoutesOperation(pickedData);
                        ApplyResult(dialog, result);
                    }
                }
                catch (Exception ex)
                {
                    ShowError(dialog, ex.ToString());
                }
            }

            private static void ApplyResult(EvacuationRoutesDialog dialog, EvacuationRoutesOperationResult result)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.ApplyOperationResult(result)));
            }

            private static void Notify(EvacuationRoutesDialog dialog, string text)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.NotifyRequestStatus(text)));
            }

            private static void Finish(EvacuationRoutesDialog dialog, string text)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.FinishRequest(text)));
            }

            private static void ShowError(EvacuationRoutesDialog dialog, string text)
            {
                if (dialog == null)
                {
                    TaskDialog.Show("Ошибка", text);
                    return;
                }

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.ShowRequestError(text)));
            }

            private static void SelectRow(EvacuationRoutesDialog dialog, long elementId)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.Invoke(new Action(() => dialog.SelectRowByElementId(elementId)));
            }

            private static void HideForPick(EvacuationRoutesDialog dialog)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.Invoke(new Action(dialog.HideForPick));
            }

            private static void RestoreAfterPick(EvacuationRoutesDialog dialog)
            {
                if (dialog == null)
                    return;

                dialog.Dispatcher.Invoke(new Action(dialog.RestoreAfterPick));
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            try
            {
                CreateEvacuationRoutes();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private sealed class StairsSelectionFilter : ISelectionFilter
        {
            private readonly Document _doc;
            public StairsSelectionFilter(Document doc) { _doc = doc; }

            public bool AllowElement(Element elem) => elem is Stairs || elem is MultistoryStairs;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        // =========================
        // ДАННЫЕ ДЛЯ БЛОКОВ МАРША
        // =========================
        private sealed class RunRouteBodyInfo
        {
            public ElementId RunId;
            public int StairsId;

            public double WidthFt;
            public double HeightFt;

            // ОСИ МАРША
            public XYZ XDirPlan; // Вдоль марша
            public XYZ YDirPlan; // Поперёк марша
            public EndFace BottomEnd; // Торец у нижней точки марша (bottomCenter)
            public EndFace TopEnd;    // Торец у верхней точки марша (topCenter)

            public IEnumerable<XYZ> GetAll8Corners()
            {
                yield return BottomEnd.BL;
                yield return BottomEnd.BR;
                yield return BottomEnd.TR;
                yield return BottomEnd.TL;
                yield return TopEnd.BL;
                yield return TopEnd.BR;
                yield return TopEnd.TR;
                yield return TopEnd.TL;
            }
        }

        private sealed class RouteBuildTarget
        {
            public Stairs Stairs;
            public long OwnerElementId;
            public long StandardStairsId;
            public ElementId PlacementLevelId;
            public double VerticalOffsetFt;
            public bool IsMultistoryPlacement;
            public string ShapeKeyPrefix;
            public string DisplayName;
        }

        private sealed class RouteBuildTargetResult
        {
            public RouteBuildTarget Target;
            public bool Ok;
            public int CreatedRuns;
            public int CreatedLandings;
            public List<int> FailedRuns = new List<int>();
            public List<int> FailedLandings = new List<int>();
            public int Intersections;
            public EvacuationRoutesStatusUpdate Update;
        }

        private sealed class RunTopSearchContext
        {
            public List<Solid> RunSolids = new List<Solid>();
            public List<Solid> FinishSolids = new List<Solid>();
            public double MinZ;
            public double MaxZ;
        }

        private struct RunClearWidthInfo
        {
            public double WidthFt;
            public double CenterOffsetFt;
        }

        private sealed class RouteIntersectionReportItem
        {
            public string RouteName;
            public long RouteElementId;
            public long OwnerElementId;
            public long ComponentElementId;
            public string ComponentKind;
            public ElementId PlacementLevelId;
            public List<RouteIntersectionTarget> Targets = new List<RouteIntersectionTarget>();
        }

        private sealed class RouteIntersectionTarget
        {
            public long ElementId;
            public string SourceName;
            public long? LinkInstanceId;
            public string CategoryName;
            public string ElementName;
        }

        private sealed class RouteIntersectionViewItem
        {
            public string RouteName { get; set; }
            public long RouteElementId { get; set; }
            public string SourceName { get; set; }
            public long? LinkInstanceId { get; set; }
            public long ElementId { get; set; }
            public string CategoryName { get; set; }
            public string ElementName { get; set; }
        }

        private sealed class RouteDebugLog
        {
            public bool Enabled;
            public List<string> Lines = new List<string>();

            public void Add(string text)
            {
                if (!Enabled) return;
                Lines.Add(text ?? "");
            }

            public void AddBlank()
            {
                if (!Enabled) return;
                Lines.Add("");
            }
        }

        private sealed class RouteIntersectionReportWindow : System.Windows.Window
        {
            private readonly UIDocument _uidoc;
            private readonly List<RouteIntersectionReportItem> _reports;
            private readonly RouteDebugLog _debugLog;
            private readonly List<RouteIntersectionViewItem> _items;
            private readonly System.Windows.Controls.ListView _list;
            private readonly System.Windows.Controls.TextBlock _status;

            public RouteIntersectionReportWindow(UIDocument uidoc, List<RouteIntersectionReportItem> reports, RouteDebugLog debugLog)
            {
                _uidoc = uidoc;
                _reports = reports ?? new List<RouteIntersectionReportItem>();
                _debugLog = debugLog;
                _items = BuildViewItems(_reports);

                Title = "Проверить пересечения с элементами";
                Width = 980;
                Height = 560;
                MinWidth = 760;
                MinHeight = 420;
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
                ResizeMode = System.Windows.ResizeMode.CanResize;

                var root = new System.Windows.Controls.Grid
                {
                    Margin = new System.Windows.Thickness(12)
                };
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
                root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

                int targetCount = _items.Count;
                int routeCount = _reports.Count(x => x != null && x.Targets != null && x.Targets.Count > 0);
                var header = new System.Windows.Controls.TextBlock
                {
                    Text = targetCount == 0
                        ? "Пересечений с элементами не найдено."
                        : $"Найдены пересечения с элементами: маршрутов {routeCount}, элементов {targetCount}. Клик по строке выбирает элемент.",
                    Margin = new System.Windows.Thickness(0, 0, 0, 8),
                    TextWrapping = System.Windows.TextWrapping.Wrap
                };
                System.Windows.Controls.Grid.SetRow(header, 0);
                root.Children.Add(header);

                _list = new System.Windows.Controls.ListView
                {
                    ItemsSource = _items,
                    HorizontalContentAlignment = System.Windows.HorizontalAlignment.Stretch,
                    Margin = new System.Windows.Thickness(0, 0, 0, 10)
                };
                _list.SelectionChanged += (sender, args) => SelectSelectedTarget();
                _list.MouseDoubleClick += (sender, args) => SelectSelectedTarget();

                var view = new System.Windows.Controls.GridView();
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "Путь",
                    Width = 220,
                    DisplayMemberBinding = new System.Windows.Data.Binding("RouteName")
                });
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "ID пути",
                    Width = 80,
                    DisplayMemberBinding = new System.Windows.Data.Binding("RouteElementId")
                });
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "Источник",
                    Width = 180,
                    DisplayMemberBinding = new System.Windows.Data.Binding("SourceName")
                });
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "ID элемента",
                    Width = 90,
                    DisplayMemberBinding = new System.Windows.Data.Binding("ElementId")
                });
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "Категория",
                    Width = 120,
                    DisplayMemberBinding = new System.Windows.Data.Binding("CategoryName")
                });
                view.Columns.Add(new System.Windows.Controls.GridViewColumn
                {
                    Header = "Тип",
                    Width = 220,
                    DisplayMemberBinding = new System.Windows.Data.Binding("ElementName")
                });
                _list.View = view;

                System.Windows.Controls.Grid.SetRow(_list, 1);
                root.Children.Add(_list);

                _status = new System.Windows.Controls.TextBlock
                {
                    Text = _items.Count == 0 ? "Нет найденных пересечений с элементами." : "Выберите строку для перехода к элементу.",
                    Margin = new System.Windows.Thickness(0, 0, 0, 8),
                    TextWrapping = System.Windows.TextWrapping.Wrap
                };
                System.Windows.Controls.Grid.SetRow(_status, 2);
                root.Children.Add(_status);

                var buttons = new System.Windows.Controls.StackPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                };

                buttons.Children.Add(CreateButton("Сохранить отчёт о пересечениях", 220, SaveIntersectionReport));
                buttons.Children.Add(CreateButton("Сохранить отчёт о марше", 180, SaveRunReport));

                System.Windows.Controls.Grid.SetRow(buttons, 3);
                root.Children.Add(buttons);

                Content = root;
            }

            private static List<RouteIntersectionViewItem> BuildViewItems(List<RouteIntersectionReportItem> reports)
            {
                var result = new List<RouteIntersectionViewItem>();
                foreach (var report in reports ?? new List<RouteIntersectionReportItem>())
                {
                    if (report == null || report.Targets == null)
                        continue;

                    foreach (var target in report.Targets.OrderBy(x => x.SourceName).ThenBy(x => x.ElementId))
                    {
                        if (target == null)
                            continue;

                        result.Add(new RouteIntersectionViewItem
                        {
                            RouteName = string.IsNullOrWhiteSpace(report.RouteName) ? "Путь эвакуации" : report.RouteName,
                            RouteElementId = report.RouteElementId,
                            SourceName = string.IsNullOrWhiteSpace(target.SourceName) ? "Host" : target.SourceName,
                            LinkInstanceId = target.LinkInstanceId,
                            ElementId = target.ElementId,
                            CategoryName = target.CategoryName,
                            ElementName = target.ElementName
                        });
                    }
                }

                return result;
            }

            private static System.Windows.Controls.Button CreateButton(string text, double width, Action action)
            {
                var button = new System.Windows.Controls.Button
                {
                    Content = text,
                    Width = width,
                    Height = 28,
                    Margin = new System.Windows.Thickness(6, 0, 0, 0)
                };
                button.Click += (sender, args) => action?.Invoke();
                return button;
            }

            private RouteIntersectionViewItem GetSelectedItem()
            {
                return _list == null ? null : _list.SelectedItem as RouteIntersectionViewItem;
            }

            private void SelectSelectedTarget()
            {
                var item = GetSelectedItem();
                if (item == null)
                {
                    SetStatus("Сначала выберите строку отчёта.");
                    return;
                }

                if (_uidoc == null || _uidoc.Document == null)
                {
                    SetStatus("Не удалось получить активный документ Revit.");
                    return;
                }

                if (item.LinkInstanceId.HasValue && item.LinkInstanceId.Value > 0)
                {
                    ElementId linkId = IDHelper.CreateElementId(item.LinkInstanceId.Value);
                    Element link = _uidoc.Document.GetElement(linkId);
                    if (link == null)
                    {
                        SetStatus($"Связь ID {item.LinkInstanceId.Value} не найдена в активном документе. Элемент внутри связи: ID {item.ElementId}.");
                        return;
                    }

                    TrySelectAndShow(linkId);
                    SetStatus($"Выбрана связь ID {item.LinkInstanceId.Value}. Элемент внутри связи: ID {item.ElementId}.");
                    return;
                }

                ElementId id = IDHelper.CreateElementId(item.ElementId);
                Element elem = _uidoc.Document.GetElement(id);
                if (elem == null)
                {
                    SetStatus($"Элемент ID {item.ElementId} не найден в активном документе.");
                    return;
                }

                TrySelectAndShow(id);
                SetStatus($"Выбран элемент ID {item.ElementId}.");
            }

            private void TrySelectAndShow(ElementId id)
            {
                try
                {
                    _uidoc.Selection.SetElementIds(new List<ElementId> { id });
                }
                catch (Exception ex)
                {
                    SetStatus($"Не удалось выбрать элемент: {ex.Message}");
                    return;
                }

                try
                {
                    _uidoc.ShowElements(id);
                }
                catch
                {
                }
            }

            private void SaveIntersectionReport()
            {
                try
                {
                    string path = SaveIntersectionReportToDesktop(_reports);
                    SetStatus($"Отчёт о пересечениях сохранён: {path}");
                }
                catch (Exception ex)
                {
                    SetStatus($"Не удалось сохранить отчёт о пересечениях: {ex.Message}");
                }
            }

            private void SaveRunReport()
            {
                try
                {
                    string path = SaveDebugLogToDesktop(_debugLog);
                    SetStatus($"Отчёт о марше сохранён: {path}");
                }
                catch (Exception ex)
                {
                    SetStatus($"Не удалось сохранить отчёт о марше: {ex.Message}");
                }
            }

            private void SetStatus(string text)
            {
                if (_status != null)
                    _status.Text = text ?? "";
            }
        }

        private struct EndFace
        {
            public XYZ BL;  // Bottom-Left
            public XYZ BR;  // Bottom-Right
            public XYZ TR;  // Top-Right
            public XYZ TL;  // Top-Left

            public XYZ Center => (BL + BR + TR + TL) * 0.25;

            public double MinZBottom => Math.Min(BL.Z, BR.Z);

            public void GetSpanOnDir(XYZ dirPlan, out double min, out double max)
            {
                min = double.PositiveInfinity;
                max = double.NegativeInfinity;

                XYZ[] pts = new[] { BL, BR, TR, TL };
                foreach (var p in pts)
                {
                    XYZ pxy = new XYZ(p.X, p.Y, 0);
                    double t = pxy.DotProduct(dirPlan);
                    if (t < min) min = t;
                    if (t > max) max = t;
                }
            }
        }

        // =========================
        // ОСНОВНАЯ ЛОГИКА
        // =========================
        private void CreateEvacuationRoutes()
        {
            var stairRows = GetStairListItems(doc);
            var evacuationWorksets = GetEvacuationWorksetOptions(doc);
            var handler = new EvacuationRoutesExternalEventHandler(this);
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            EvacuationRoutesDialog dlg = null;
            dlg = new EvacuationRoutesDialog(
                stairRows,
                evacuationWorksets,
                data =>
                {
                    handler.RequestPickAndBuild(dlg, data);
                    externalEvent.Raise();
                },
                id =>
                {
                    handler.RequestSelect(dlg, id);
                    externalEvent.Raise();
                },
                data =>
                {
                    handler.RequestBuild(dlg, data);
                    externalEvent.Raise();
                });

            new WindowInteropHelper(dlg) { Owner = uiapp.MainWindowHandle };
            dlg.Show();
        }

        private List<EvacuationRoutesStairListItem> GetStairListItems(Document doc)
        {
            var result = new List<EvacuationRoutesStairListItem>();
            var seenStairs = new HashSet<long>();
            if (doc == null)
                return result;

            foreach (MultistoryStairs multistory in new FilteredElementCollector(doc)
                .OfClass(typeof(MultistoryStairs))
                .WhereElementIsNotElementType()
                .OfType<MultistoryStairs>())
            {
                Stairs standardStairs = GetMultistoryStandardStairs(doc, multistory);
                var nestedIds = GetMultistoryStandardStairIds(doc, multistory)
                    .Select(x => IDHelper.ElIdValue(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();
                if (standardStairs != null)
                {
                    long standardStairsId = IDHelper.ElIdValue(standardStairs.Id);
                    if (standardStairsId > 0 && !nestedIds.Contains(standardStairsId))
                        nestedIds.Add(standardStairsId);
                }

                int connectedLevelCount = GetMultistoryConnectedLevelIds(multistory).Count;
                int placementCount = GetMultistoryPlacementLevelIds(doc, multistory, standardStairs).Count;

                result.Add(new EvacuationRoutesStairListItem
                {
                    ElementId = IDHelper.ElIdValue(multistory.Id),
                    Kind = "Многоэтажная",
                    Name = GetElementDisplayName(multistory),
                    TypeName = GetElementTypeDisplayName(doc, multistory),
                    RunCount = GetStairRunCount(standardStairs),
                    LandingCount = GetStairLandingCount(standardStairs),
                    NestedCount = placementCount > 0 ? placementCount : nestedIds.Count,
                    ConnectedLevelCount = connectedLevelCount,
                    NestedStairIds = nestedIds
                });

                foreach (long nestedId in nestedIds)
                    seenStairs.Add(nestedId);
            }

            foreach (Stairs stairs in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .OfType<Stairs>())
            {
                AddStairListItemIfNew(doc, stairs, result, seenStairs, null);
            }

            return result
                .OrderBy(GetStairListGroupKey)
                .ThenBy(x => x.ParentMultistoryId.HasValue ? 1 : 0)
                .ThenBy(x => x.ElementId)
                .ToList();
        }

        private static long GetStairListGroupKey(EvacuationRoutesStairListItem item)
        {
            if (item == null)
                return long.MaxValue;

            return item.ParentMultistoryId ?? item.ElementId;
        }

        private static void AddStairListItemIfNew(Document doc, Stairs stairs, List<EvacuationRoutesStairListItem> result, HashSet<long> seenStairs, long? parentMultistoryId)
        {
            if (doc == null || stairs == null || result == null || seenStairs == null)
                return;

            long id = IDHelper.ElIdValue(stairs.Id);
            if (!seenStairs.Add(id))
                return;

            var item = new EvacuationRoutesStairListItem
            {
                ElementId = id,
                Kind = parentMultistoryId.HasValue ? "Стандартная" : "Лестница",
                Name = GetElementDisplayName(stairs),
                TypeName = GetElementTypeDisplayName(doc, stairs),
                RunCount = GetStairRunCount(stairs),
                LandingCount = GetStairLandingCount(stairs),
                NestedCount = 0,
                ParentMultistoryId = parentMultistoryId
            };

            result.Add(item);
        }

        private static int GetStairRunCount(Stairs stairs)
        {
            try
            {
                ICollection<ElementId> ids = stairs?.GetStairsRuns();
                return ids == null ? 0 : ids.Count;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetStairLandingCount(Stairs stairs)
        {
            try
            {
                ICollection<ElementId> ids = stairs?.GetStairsLandings();
                return ids == null ? 0 : ids.Count;
            }
            catch
            {
                return 0;
            }
        }

        private EvacuationRoutesOperationResult RunEvacuationRoutesOperation(EvacuationRoutesDialogResult data)
        {
            var result = new EvacuationRoutesOperationResult();
            if (data == null)
                return result;

            List<RouteBuildTarget> targets = GetTargetRouteBuildTargets(doc, data);
            if (targets == null || targets.Count == 0)
                throw new InvalidOperationException(data.PickSingleStair ? "Для выбранной строки не найдены лестницы для обработки." : "В документе не найдены лестницы для обработки.");

            var debugLog = new RouteDebugLog { Enabled = data.PickSingleStair };
            if (debugLog.Enabled)
            {
                debugLog.Add("KPLN. Пути эвакуации — DEBUG-отчёт одиночного запуска");
                debugLog.Add($"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                debugLog.Add($"Документ: {doc.Title}");
                debugLog.Add($"Настройки: HeightMm={data.HeightMm}; WidthMm={data.WidthMm}; UseRunWidth={data.UseRunWidth}; AddToEvacuationWorkset={data.AddToEvacuationWorkset}; WorksetId={(data.EvacuationWorksetId.HasValue ? data.EvacuationWorksetId.Value.ToString() : "null")}; SelectedElementId={(data.SelectedElementId.HasValue ? data.SelectedElementId.Value.ToString() : "null")}");
                debugLog.AddBlank();

                foreach (RouteBuildTarget target in targets)
                    AddRouteBuildTargetDebugHeader(debugLog, target);
            }

            var buildResults = new List<RouteBuildTargetResult>();
            var intersectionReports = new List<RouteIntersectionReportItem>();

            using (var t = new Transaction(doc, "KPLN: Построение путей эвакуации"))
            {
                t.Start();

                foreach (RouteBuildTarget target in targets)
                {
                    Stairs stairs = target?.Stairs;
                    if (stairs == null)
                        continue;

                    long stairId = IDHelper.ElIdValue(stairs.Id);
                    int beforeIntersections = intersectionReports.Count;

                    int stairCreatedRuns;
                    int stairCreatedLandings;
                    List<int> stairFailedRuns;
                    List<int> stairFailedLandings;

                    bool okStair = TryCreateRouteBodyOnStair(
                        doc, target, data, intersectionReports, debugLog,
                        out stairCreatedRuns, out stairCreatedLandings,
                        out stairFailedRuns, out stairFailedLandings);

                    int routeIntersections = intersectionReports
                        .Skip(beforeIntersections)
                        .Sum(x => x.Targets == null ? 0 : x.Targets.Count);

                    EvacuationRoutesStatusUpdate update = CreateStairStatusUpdate(
                        stairId,
                        okStair,
                        stairCreatedRuns,
                        stairCreatedLandings,
                        stairFailedRuns,
                        stairFailedLandings,
                        routeIntersections);

                    var buildResult = new RouteBuildTargetResult
                    {
                        Target = target,
                        Ok = okStair,
                        CreatedRuns = stairCreatedRuns,
                        CreatedLandings = stairCreatedLandings,
                        FailedRuns = stairFailedRuns ?? new List<int>(),
                        FailedLandings = stairFailedLandings ?? new List<int>(),
                        Intersections = routeIntersections,
                        Update = update
                    };

                    buildResults.Add(buildResult);

                    if (!target.IsMultistoryPlacement)
                        result.Updates.Add(update);
                }

                t.Commit();
            }

            AddMultistoryAggregateUpdates(result, buildResults);
            FillOperationReportLines(result, data, targets, intersectionReports, debugLog, buildResults);

            return result;
        }

        private List<RouteBuildTarget> GetTargetRouteBuildTargets(Document doc, EvacuationRoutesDialogResult data)
        {
            if (doc == null || data == null)
                return new List<RouteBuildTarget>();

            if (!data.PickSingleStair)
                return FilterRouteBuildTargetsByIncludedIds(GetRouteBuildTargets(doc), data.IncludedElementIds);

            if (!data.SelectedElementId.HasValue)
                return new List<RouteBuildTarget>();

            Element selected = doc.GetElement(IDHelper.CreateElementId(data.SelectedElementId.Value));

            MultistoryStairs multistory = selected as MultistoryStairs;
            if (multistory != null)
                return GetMultistoryRouteBuildTargets(doc, multistory);

            Stairs stairs = selected as Stairs;
            if (stairs == null)
                return new List<RouteBuildTarget>();

            MultistoryStairs parent = GetParentMultistoryStairs(doc, stairs);
            if (parent != null)
                return GetMultistoryRouteBuildTargets(doc, parent);

            return new List<RouteBuildTarget> { CreateOrdinaryRouteBuildTarget(stairs) };
        }

        private static List<RouteBuildTarget> FilterRouteBuildTargetsByIncludedIds(List<RouteBuildTarget> targets, IEnumerable<long> includedElementIds)
        {
            if (targets == null)
                return new List<RouteBuildTarget>();

            if (includedElementIds == null)
                return targets;

            var included = new HashSet<long>(includedElementIds.Where(x => x > 0));
            if (included.Count == 0)
                return new List<RouteBuildTarget>();

            return targets
                .Where(x => x != null && (included.Contains(x.OwnerElementId) || (!x.IsMultistoryPlacement && included.Contains(x.StandardStairsId))))
                .ToList();
        }

        private static List<RouteBuildTarget> GetRouteBuildTargets(Document doc)
        {
            var result = new List<RouteBuildTarget>();
            var multistoryStandardIds = new HashSet<long>();
            if (doc == null)
                return result;

            foreach (MultistoryStairs multistory in new FilteredElementCollector(doc)
                .OfClass(typeof(MultistoryStairs))
                .WhereElementIsNotElementType()
                .OfType<MultistoryStairs>())
            {
                var targets = GetMultistoryRouteBuildTargets(doc, multistory);
                result.AddRange(targets);

                foreach (RouteBuildTarget target in targets)
                    multistoryStandardIds.Add(target.StandardStairsId);
            }

            foreach (Stairs stairs in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .OfType<Stairs>())
            {
                long id = IDHelper.ElIdValue(stairs.Id);
                if (multistoryStandardIds.Contains(id))
                    continue;

                result.Add(CreateOrdinaryRouteBuildTarget(stairs));
            }

            return result
                .Where(x => x != null && x.Stairs != null)
                .OrderBy(x => x.OwnerElementId)
                .ThenBy(x => x.IsMultistoryPlacement ? GetLevelSortKey(doc, x.PlacementLevelId) : 0.0)
                .ThenBy(x => x.StandardStairsId)
                .ToList();
        }

        private static List<RouteBuildTarget> GetMultistoryRouteBuildTargets(Document doc, MultistoryStairs multistory)
        {
            var result = new List<RouteBuildTarget>();
            if (doc == null || multistory == null)
                return result;

            Stairs standardStairs = GetMultistoryStandardStairs(doc, multistory);
            if (standardStairs == null)
                return result;

            ElementId standardTopLevelId;
            GetStairsLevelIds(standardStairs, out ElementId standardBaseLevelId, out standardTopLevelId);
            double standardBaseElevation = GetLevelElevation(doc, standardBaseLevelId) ?? 0.0;
            long multistoryId = IDHelper.ElIdValue(multistory.Id);
            long standardStairsId = IDHelper.ElIdValue(standardStairs.Id);

            List<ElementId> placementLevels = GetMultistoryPlacementLevelIds(doc, multistory, standardStairs);
            if (placementLevels.Count == 0)
            {
                result.Add(CreateMultistoryRouteBuildTarget(doc, multistoryId, standardStairs, standardBaseLevelId, 0.0));
                return result;
            }

            foreach (ElementId levelId in placementLevels)
            {
                double levelElevation = GetLevelElevation(doc, levelId) ?? standardBaseElevation;
                result.Add(CreateMultistoryRouteBuildTarget(doc, multistoryId, standardStairs, levelId, levelElevation - standardBaseElevation));
            }

            return result
                .Where(x => x != null && x.Stairs != null)
                .GroupBy(x => $"{x.StandardStairsId}:{IDHelper.ElIdValue(x.PlacementLevelId)}")
                .Select(x => x.First())
                .OrderBy(x => GetLevelSortKey(doc, x.PlacementLevelId))
                .ThenBy(x => standardStairsId)
                .ToList();
        }

        private static RouteBuildTarget CreateOrdinaryRouteBuildTarget(Stairs stairs)
        {
            if (stairs == null)
                return null;

            long stairId = IDHelper.ElIdValue(stairs.Id);
            return new RouteBuildTarget
            {
                Stairs = stairs,
                OwnerElementId = stairId,
                StandardStairsId = stairId,
                PlacementLevelId = ElementId.InvalidElementId,
                VerticalOffsetFt = 0.0,
                IsMultistoryPlacement = false,
                ShapeKeyPrefix = "",
                DisplayName = $"Лестница {stairId}"
            };
        }

        private static RouteBuildTarget CreateMultistoryRouteBuildTarget(Document doc, long multistoryId, Stairs standardStairs, ElementId placementLevelId, double verticalOffsetFt)
        {
            if (standardStairs == null)
                return null;

            long standardStairsId = IDHelper.ElIdValue(standardStairs.Id);
            long levelId = placementLevelId == null || placementLevelId == ElementId.InvalidElementId ? 0 : IDHelper.ElIdValue(placementLevelId);
            string levelName = "";
            try { levelName = doc?.GetElement(placementLevelId)?.Name ?? ""; } catch { }

            return new RouteBuildTarget
            {
                Stairs = standardStairs,
                OwnerElementId = multistoryId,
                StandardStairsId = standardStairsId,
                PlacementLevelId = placementLevelId ?? ElementId.InvalidElementId,
                VerticalOffsetFt = verticalOffsetFt,
                IsMultistoryPlacement = true,
                ShapeKeyPrefix = $"MS_{multistoryId}_{levelId}",
                DisplayName = $"Многоэтажная {multistoryId}; стандартная {standardStairsId}; уровень {levelId} {levelName}".Trim()
            };
        }

        private static double GetLevelSortKey(Document doc, ElementId levelId)
        {
            return GetLevelElevation(doc, levelId) ?? double.MaxValue;
        }

        private static double? GetLevelElevation(Document doc, ElementId levelId)
        {
            if (doc == null || levelId == null || levelId == ElementId.InvalidElementId)
                return null;

            try
            {
                Level level = doc.GetElement(levelId) as Level;
                return level?.Elevation;
            }
            catch
            {
                return null;
            }
        }

        private static EvacuationRoutesStatusUpdate CreateStairStatusUpdate(
            long stairId,
            bool okStair,
            int createdRuns,
            int createdLandings,
            List<int> failedRuns,
            List<int> failedLandings,
            int intersections)
        {
            failedRuns = failedRuns ?? new List<int>();
            failedLandings = failedLandings ?? new List<int>();

            var details = new List<string>
            {
                $"Создано маршей: {createdRuns}",
                $"Создано площадок: {createdLandings}"
            };

            if (failedRuns.Count > 0)
                details.Add(FormatIdsLine("Необработанные марши", failedRuns));

            if (failedLandings.Count > 0)
                details.Add(FormatIdsLine("Необработанные площадки", failedLandings));

            if (intersections > 0)
                details.Add($"Пересечения с элементами: {intersections}");

            if (!okStair)
            {
                details.Add("Лестница не построилась.");
                return new EvacuationRoutesStatusUpdate
                {
                    ElementId = stairId,
                    Status = EvacuationRoutesStatus.Error,
                    StatusText = "Не построено",
                    Message = string.Join("; ", details)
                };
            }

            if (intersections > 0)
            {
                return new EvacuationRoutesStatusUpdate
                {
                    ElementId = stairId,
                    Status = EvacuationRoutesStatus.Warning,
                    StatusText = "Проблемы",
                    Message = string.Join("; ", details)
                };
            }

            if (failedRuns.Count > 0 || failedLandings.Count > 0)
            {
                return new EvacuationRoutesStatusUpdate
                {
                    ElementId = stairId,
                    Status = EvacuationRoutesStatus.Warning,
                    StatusText = "Проблемы",
                    Message = string.Join("; ", details)
                };
            }

            return new EvacuationRoutesStatusUpdate
            {
                ElementId = stairId,
                Status = EvacuationRoutesStatus.Ok,
                StatusText = "ОК",
                Message = string.Join("; ", details)
            };
        }

        private void AddMultistoryAggregateUpdates(EvacuationRoutesOperationResult result, List<RouteBuildTargetResult> buildResults)
        {
            if (result == null || buildResults == null || buildResults.Count == 0)
                return;

            foreach (var group in buildResults
                .Where(x => x != null && x.Target != null && x.Target.IsMultistoryPlacement)
                .GroupBy(x => x.Target.OwnerElementId))
            {
                var placementResults = group
                    .Where(x => x.Update != null)
                    .ToList();

                if (placementResults.Count == 0)
                    continue;

                EvacuationRoutesStatus status;
                string text;
                if (placementResults.All(x => x.Update.Status == EvacuationRoutesStatus.Error))
                {
                    status = EvacuationRoutesStatus.Error;
                    text = "Не построено";
                }
                else if (placementResults.Any(x => x.Update.Status == EvacuationRoutesStatus.Error || x.Update.Status == EvacuationRoutesStatus.Warning))
                {
                    status = EvacuationRoutesStatus.Warning;
                    text = "Проблемы";
                }
                else
                {
                    status = EvacuationRoutesStatus.Ok;
                    text = "ОК";
                }

                int createdRuns = placementResults.Sum(x => x.CreatedRuns);
                int createdLandings = placementResults.Sum(x => x.CreatedLandings);
                int intersections = placementResults.Sum(x => x.Intersections);
                int red = placementResults.Count(x => x.Update.Status == EvacuationRoutesStatus.Error);
                int yellow = placementResults.Count(x => x.Update.Status == EvacuationRoutesStatus.Warning);
                long standardStairsId = placementResults.Select(x => x.Target.StandardStairsId).FirstOrDefault();

                result.Updates.Add(new EvacuationRoutesStatusUpdate
                {
                    ElementId = group.Key,
                    Status = status,
                    StatusText = text,
                    Message = $"Многоэтажная лестница. Стандартная лестница: {standardStairsId}; размещений: {placementResults.Count}; создано маршей: {createdRuns}; создано площадок: {createdLandings}; пересечения с элементами: {intersections}; красных: {red}; жёлтых: {yellow}."
                });

            }
        }

        private void FillOperationReportLines(EvacuationRoutesOperationResult result, EvacuationRoutesDialogResult data, List<RouteBuildTarget> targets, List<RouteIntersectionReportItem> intersectionReports, RouteDebugLog debugLog, List<RouteBuildTargetResult> buildResults)
        {
            if (result == null)
                return;

            result.ReportLines.AddRange(BuildProblemReportLines(buildResults, intersectionReports));
        }

        private static List<string> BuildProblemReportLines(List<RouteBuildTargetResult> buildResults, List<RouteIntersectionReportItem> intersectionReports)
        {
            var grouped = new SortedDictionary<long, List<string>>();

            Action<long, string> add = (stairId, line) =>
            {
                if (string.IsNullOrWhiteSpace(line))
                    return;

                if (!grouped.TryGetValue(stairId, out List<string> lines))
                {
                    lines = new List<string>();
                    grouped[stairId] = lines;
                }

                if (!lines.Contains(line))
                    lines.Add(line);
            };

            foreach (RouteBuildTargetResult buildResult in buildResults ?? new List<RouteBuildTargetResult>())
            {
                RouteBuildTarget target = buildResult?.Target;
                if (target == null)
                    continue;

                long stairId = target.OwnerElementId > 0 ? target.OwnerElementId : target.StandardStairsId;

                foreach (int runId in buildResult.FailedRuns ?? new List<int>())
                    add(stairId, $"- {FormatRouteComponentLabel(target, "Марш", runId)} - не построился");

                foreach (int landingId in buildResult.FailedLandings ?? new List<int>())
                    add(stairId, $"- {FormatRouteComponentLabel(target, "Площадка", landingId)} - не построилась");

                if (!buildResult.Ok && (buildResult.FailedRuns == null || buildResult.FailedRuns.Count == 0) && (buildResult.FailedLandings == null || buildResult.FailedLandings.Count == 0))
                    add(stairId, $"- {FormatRoutePlacementLabel(target)} - не построилось");
            }

            foreach (RouteIntersectionReportItem report in intersectionReports ?? new List<RouteIntersectionReportItem>())
            {
                if (report == null || report.Targets == null || report.Targets.Count == 0)
                    continue;

                long stairId = report.OwnerElementId > 0 ? report.OwnerElementId : 0;
                string component = FormatRouteComponentLabel(report.ComponentKind, report.ComponentElementId, report.PlacementLevelId);

                foreach (RouteIntersectionTarget target in report.Targets.OrderBy(x => x.ElementId))
                    add(stairId, $"- {component} - пересечение: {FormatIntersectionTargetShort(target)}");
            }

            var result = new List<string>();
            if (grouped.Count == 0)
            {
                result.Add("Ошибок и пересечений не найдено.");
                return result;
            }

            foreach (var group in grouped)
            {
                result.Add($"ЛЕСТНИЦА ID {group.Key}");
                result.AddRange(group.Value);
                result.Add("");
            }

            if (result.Count > 0 && string.IsNullOrWhiteSpace(result[result.Count - 1]))
                result.RemoveAt(result.Count - 1);

            return result;
        }

        private static string FormatRouteComponentLabel(RouteBuildTarget target, string kind, long componentId)
        {
            return FormatRouteComponentLabel(kind, componentId, target == null ? ElementId.InvalidElementId : target.PlacementLevelId);
        }

        private static string FormatRouteComponentLabel(string kind, long componentId, ElementId placementLevelId)
        {
            string label = string.IsNullOrWhiteSpace(kind) ? "Элемент" : kind;
            string id = componentId > 0 ? $" ID {componentId}" : "";
            string level = placementLevelId != null && placementLevelId != ElementId.InvalidElementId
                ? $" | уровень {FormatOptionalElementId(placementLevelId)}"
                : "";

            return $"{label}{id}{level}";
        }

        private static string FormatRoutePlacementLabel(RouteBuildTarget target)
        {
            if (target == null)
                return "Лестница";

            if (target.IsMultistoryPlacement)
                return $"Размещение уровня {FormatOptionalElementId(target.PlacementLevelId)}";

            long id = target.OwnerElementId > 0 ? target.OwnerElementId : target.StandardStairsId;
            return id > 0 ? $"Лестница ID {id}" : "Лестница";
        }

        private static string FormatIntersectionTargetShort(RouteIntersectionTarget target)
        {
            if (target == null)
                return "элемент";

            string source = string.IsNullOrWhiteSpace(target.SourceName) ? "Host" : target.SourceName;
            string link = target.LinkInstanceId.HasValue ? $" | LinkInstanceId {target.LinkInstanceId.Value}" : "";
            string cat = string.IsNullOrWhiteSpace(target.CategoryName) ? "без категории" : target.CategoryName;
            string name = string.IsNullOrWhiteSpace(target.ElementName) ? "" : $" | {target.ElementName}";

            return $"{source}{link} | ID {target.ElementId} | {cat}{name}";
        }

        private static void AddMultistoryDiagnostics(List<string> lines, Document doc)
        {
            if (lines == null)
                return;

            lines.Add("Диагностика многоэтажных лестниц:");

            if (doc == null)
            {
                lines.Add("Документ недоступен.");
                return;
            }

            List<MultistoryStairs> multistories;
            try
            {
                multistories = new FilteredElementCollector(doc)
                    .OfClass(typeof(MultistoryStairs))
                    .WhereElementIsNotElementType()
                    .OfType<MultistoryStairs>()
                    .OrderBy(x => IDHelper.ElIdValue(x.Id))
                    .ToList();
            }
            catch (Exception ex)
            {
                lines.Add($"Ошибка коллектора MultistoryStairs: {ex.Message}");
                return;
            }

            lines.Add($"Контейнеров MultistoryStairs: {multistories.Count}");
            if (multistories.Count == 0)
                return;

            foreach (MultistoryStairs multistory in multistories)
            {
                long multistoryId = IDHelper.ElIdValue(multistory.Id);
                List<ElementId> allStairsIds = GetMultistoryStairsIds(multistory);
                List<ElementId> resolvedMemberIds = GetMultistoryMemberStairIds(doc, multistory);
                List<ElementId> connectedLevelIds = TryInvokeElementIdCollection(multistory, "GetAllConnectedLevels");
                List<ElementId> bboxStairsIds = GetStairsIntersectingBoundingBox(doc, multistory).Select(x => x.Id).ToList();
                Stairs standardStairs = GetMultistoryStandardStairs(doc, multistory);
                List<ElementId> placementLevelIds = GetMultistoryPlacementLevelIds(doc, multistory, standardStairs);

                lines.Add("");
                lines.Add($"Multistory ID {multistoryId} | {GetElementDisplayName(multistory)} | type='{GetElementTypeDisplayName(doc, multistory)}'");
                lines.Add($"  BoundingBox={FormatBoundingBox(SafeGetBoundingBox(multistory))}");
                lines.Add($"  GetAllStairsIds count={allStairsIds.Count}: {FormatElementIds(allStairsIds)}");
                lines.Add($"  StandardStairsId={(standardStairs == null ? "нет" : IDHelper.ElIdValue(standardStairs.Id).ToString())}");
                lines.Add($"  Placement levels count={placementLevelIds.Count}: {FormatElementIds(placementLevelIds)}");
                lines.Add($"  Resolved member stairs count={resolvedMemberIds.Count}: {FormatElementIds(resolvedMemberIds)}");
                lines.Add($"  GetAllConnectedLevels count={connectedLevelIds.Count}: {FormatElementIds(connectedLevelIds)}");
                lines.Add($"  Stairs by bbox overlap count={bboxStairsIds.Count}: {FormatElementIds(bboxStairsIds)}");
                AddMultistoryCandidateDiagnostics(lines, doc, multistory, connectedLevelIds, resolvedMemberIds, bboxStairsIds);
                AddMultistoryApiDiagnostics(lines, multistory);

                if (connectedLevelIds.Count > 0)
                {
                    foreach (ElementId levelId in connectedLevelIds)
                    {
                        List<ElementId> stairsOnLevelIds = TryInvokeElementIdCollection(multistory, "GetStairsOnLevel", levelId);
                        string levelName = "";
                        try { levelName = doc.GetElement(levelId)?.Name ?? ""; } catch { }
                        lines.Add($"  Level {IDHelper.ElIdValue(levelId)} '{levelName}' -> GetStairsOnLevel count={stairsOnLevelIds.Count}: {FormatElementIds(stairsOnLevelIds)}");
                    }
                }
            }
        }

        private static void AddMultistoryCandidateDiagnostics(List<string> lines, Document doc, MultistoryStairs multistory, List<ElementId> connectedLevelIds, List<ElementId> resolvedMemberIds, List<ElementId> bboxStairsIds)
        {
            if (lines == null || doc == null || multistory == null)
                return;

            var connectedLevels = new HashSet<long>((connectedLevelIds ?? new List<ElementId>()).Select(IDHelper.ElIdValue));
            var resolvedIds = new HashSet<long>((resolvedMemberIds ?? new List<ElementId>()).Select(IDHelper.ElIdValue));
            var bboxIds = new HashSet<long>((bboxStairsIds ?? new List<ElementId>()).Select(IDHelper.ElIdValue));
            BoundingBoxXYZ multistoryBox = SafeGetBoundingBox(multistory);
            double xyToleranceFt = MmToInternal(1500.0);

            List<Stairs> candidates;
            try
            {
                candidates = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .OfType<Stairs>()
                    .OrderBy(x => IDHelper.ElIdValue(x.Id))
                    .ToList();
            }
            catch (Exception ex)
            {
                lines.Add($"  Candidate diagnostics ERROR: {ex.Message}");
                return;
            }

            var reportLines = new List<string>();
            foreach (Stairs stairs in candidates)
            {
                long id = IDHelper.ElIdValue(stairs.Id);
                GetStairsLevelIds(stairs, out ElementId baseLevelId, out ElementId topLevelId);
                bool baseConnected = IsLevelInSet(baseLevelId, connectedLevels);
                bool topConnected = IsLevelInSet(topLevelId, connectedLevels);
                bool levelPairFits = baseConnected && topConnected;
                bool resolved = resolvedIds.Contains(id);
                bool bboxOverlap = bboxIds.Contains(id);
                bool xyNearContainer = BoundingBoxesIntersectXY(multistoryBox, SafeGetBoundingBox(stairs), xyToleranceFt);
                string guid = GetElementIfcGuid(doc, stairs);

                if (!resolved && !levelPairFits && !bboxOverlap && !xyNearContainer)
                    continue;

                reportLines.Add(
                    $"    ID {id} | base={FormatOptionalElementId(baseLevelId)} | top={FormatOptionalElementId(topLevelId)} | levels={FormatYesNo(levelPairFits)} | resolved={FormatYesNo(resolved)} | bboxXYZ={FormatYesNo(bboxOverlap)} | bboxXY+1500={FormatYesNo(xyNearContainer)} | ifc='{guid}' | type='{GetElementTypeDisplayName(doc, stairs)}'");
            }

            lines.Add($"  Candidate stairs diagnostics count={reportLines.Count}:");
            if (reportLines.Count == 0)
                lines.Add("    нет");
            else
                lines.AddRange(reportLines);
        }

        private static void AddMultistoryApiDiagnostics(List<string> lines, MultistoryStairs multistory)
        {
            if (lines == null || multistory == null)
                return;

            try
            {
                Type type = multistory.GetType();
                var methods = type
                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(x => x.DeclaringType != typeof(object))
                    .Where(x => x.Name.IndexOf("Stair", StringComparison.OrdinalIgnoreCase) >= 0
                        || x.Name.IndexOf("Level", StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(x => x.Name)
                    .Select(FormatMethodSignature)
                    .Distinct()
                    .ToList();

                var properties = type
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .Where(x => x.Name.IndexOf("Stair", StringComparison.OrdinalIgnoreCase) >= 0
                        || x.Name.IndexOf("Level", StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(x => x.Name)
                    .Select(x => $"{x.PropertyType.Name} {x.Name}")
                    .Distinct()
                    .ToList();

                lines.Add($"  API methods Stair/Level count={methods.Count}: {(methods.Count == 0 ? "нет" : string.Join("; ", methods))}");
                lines.Add($"  API properties Stair/Level count={properties.Count}: {(properties.Count == 0 ? "нет" : string.Join("; ", properties))}");
            }
            catch (Exception ex)
            {
                lines.Add($"  API diagnostics ERROR: {ex.Message}");
            }
        }

        private static string FormatMethodSignature(MethodInfo method)
        {
            if (method == null)
                return "";

            string parameters = string.Join(", ", method.GetParameters().Select(x => $"{x.ParameterType.Name} {x.Name}"));
            return $"{method.ReturnType.Name} {method.Name}({parameters})";
        }

        private static List<ElementId> GetMultistoryMemberStairIds(Document doc, MultistoryStairs multistory)
        {
            var result = new List<ElementId>();
            if (doc == null || multistory == null)
                return result;

            AddElementIds(result, GetMultistoryStairsIds(multistory));
            AddElementIdIfValid(result, TryGetElementIdProperty(multistory, "StandardStairsId"));

            foreach (ElementId levelId in GetMultistoryConnectedLevelIds(multistory))
                AddElementIds(result, TryInvokeElementIdCollection(multistory, "GetStairsOnLevel", levelId));

            return result
                .Where(x => x != null && x != ElementId.InvalidElementId)
                .GroupBy(IDHelper.ElIdValue)
                .Select(x => x.First())
                .OrderBy(IDHelper.ElIdValue)
                .ToList();
        }

        private static List<ElementId> GetMultistoryStandardStairIds(Document doc, MultistoryStairs multistory)
        {
            var result = new List<ElementId>();
            if (doc == null || multistory == null)
                return result;

            AddElementIdIfValid(result, TryGetElementIdProperty(multistory, "StandardStairsId"));
            AddElementIds(result, GetMultistoryStairsIds(multistory));

            return result
                .Where(x => x != null && x != ElementId.InvalidElementId)
                .GroupBy(IDHelper.ElIdValue)
                .Select(x => x.First())
                .OrderBy(IDHelper.ElIdValue)
                .ToList();
        }

        private static Stairs GetMultistoryStandardStairs(Document doc, MultistoryStairs multistory)
        {
            if (doc == null || multistory == null)
                return null;

            foreach (ElementId id in GetMultistoryStandardStairIds(doc, multistory))
            {
                Stairs stairs = doc.GetElement(id) as Stairs;
                if (stairs != null)
                    return stairs;
            }

            return null;
        }

        private static List<ElementId> GetMultistoryPlacementLevelIds(Document doc, MultistoryStairs multistory, Stairs standardStairs)
        {
            var result = new List<ElementId>();
            if (doc == null || multistory == null)
                return result;

            if (standardStairs != null)
                AddElementIds(result, TryInvokeElementIdCollection(multistory, "GetStairsPlacementLevels", standardStairs));

            if (result.Count == 0 && standardStairs != null)
            {
                long standardId = IDHelper.ElIdValue(standardStairs.Id);
                foreach (ElementId levelId in GetMultistoryConnectedLevelIds(multistory))
                {
                    List<ElementId> stairsOnLevel = TryInvokeElementIdCollection(multistory, "GetStairsOnLevel", levelId);
                    if (stairsOnLevel.Any(x => IDHelper.ElIdValue(x) == standardId))
                        AddElementIdIfValid(result, levelId);
                }
            }

            if (result.Count == 0)
            {
                var connectedLevels = GetMultistoryConnectedLevelIds(multistory)
                    .OrderBy(x => GetLevelElevation(doc, x) ?? 0.0)
                    .ToList();

                for (int i = 0; i < Math.Max(0, connectedLevels.Count - 1); i++)
                    AddElementIdIfValid(result, connectedLevels[i]);
            }

            return result
                .Where(x => x != null && x != ElementId.InvalidElementId)
                .GroupBy(IDHelper.ElIdValue)
                .Select(x => x.First())
                .OrderBy(x => GetLevelElevation(doc, x) ?? 0.0)
                .ToList();
        }

        private static void AddStairsLikelyBelongingToMultistory(Document doc, MultistoryStairs multistory, List<ElementId> result)
        {
            if (doc == null || multistory == null || result == null || result.Count == 0)
                return;

            var connectedLevels = new HashSet<long>(
                GetMultistoryConnectedLevelIds(multistory).Select(IDHelper.ElIdValue));
            if (connectedLevels.Count == 0)
                return;

            var seedIds = new HashSet<long>(result.Select(IDHelper.ElIdValue));
            var seedTypeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seedGuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seedBoxes = new List<BoundingBoxXYZ>();

            foreach (ElementId id in result.ToList())
            {
                Element elem = doc.GetElement(id);
                if (elem == null)
                    continue;

                string typeName = GetElementTypeDisplayName(doc, elem);
                if (!string.IsNullOrWhiteSpace(typeName))
                    seedTypeNames.Add(typeName);

                string guid = GetElementIfcGuid(doc, elem);
                if (!string.IsNullOrWhiteSpace(guid))
                    seedGuids.Add(guid);

                BoundingBoxXYZ bb = SafeGetBoundingBox(elem);
                if (bb != null)
                    seedBoxes.Add(bb);
            }

            BoundingBoxXYZ multistoryBox = SafeGetBoundingBox(multistory);
            double xyToleranceFt = MmToInternal(1500.0);

            try
            {
                foreach (Stairs stairs in new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .OfType<Stairs>())
                {
                    if (stairs == null || seedIds.Contains(IDHelper.ElIdValue(stairs.Id)))
                        continue;

                    GetStairsLevelIds(stairs, out ElementId baseLevelId, out ElementId topLevelId);
                    bool levelPairFits = IsLevelInSet(baseLevelId, connectedLevels) && IsLevelInSet(topLevelId, connectedLevels);
                    if (!levelPairFits)
                        continue;

                    string guid = GetElementIfcGuid(doc, stairs);
                    if (!string.IsNullOrWhiteSpace(guid) && seedGuids.Contains(guid))
                    {
                        AddElementIdIfValid(result, stairs.Id);
                        seedIds.Add(IDHelper.ElIdValue(stairs.Id));
                        continue;
                    }

                    BoundingBoxXYZ stairBox = SafeGetBoundingBox(stairs);
                    bool nearKnownFootprint = seedBoxes.Any(x => BoundingBoxesIntersectXY(x, stairBox, xyToleranceFt));
                    bool nearContainerFootprint = BoundingBoxesIntersectXY(multistoryBox, stairBox, xyToleranceFt);
                    string typeName = GetElementTypeDisplayName(doc, stairs);
                    bool sameType = !string.IsNullOrWhiteSpace(typeName) && seedTypeNames.Contains(typeName);
                    if (nearKnownFootprint || nearContainerFootprint || sameType)
                    {
                        AddElementIdIfValid(result, stairs.Id);
                        seedIds.Add(IDHelper.ElIdValue(stairs.Id));
                    }
                }
            }
            catch
            {
            }
        }

        private static void AddStairsWithMatchingIfcGuid(Document doc, List<ElementId> result)
        {
            if (doc == null || result == null || result.Count == 0)
                return;

            var guids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ElementId id in result.ToList())
            {
                Element elem = doc.GetElement(id);
                string guid = GetElementIfcGuid(doc, elem);
                if (!string.IsNullOrWhiteSpace(guid))
                    guids.Add(guid);
            }

            if (guids.Count == 0)
                return;

            try
            {
                foreach (Stairs stairs in new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .OfType<Stairs>())
                {
                    string guid = GetElementIfcGuid(doc, stairs);
                    if (!string.IsNullOrWhiteSpace(guid) && guids.Contains(guid))
                        AddElementIdIfValid(result, stairs.Id);
                }
            }
            catch
            {
            }
        }

        private static string GetElementIfcGuid(Document doc, Element elem)
        {
            if (elem == null)
                return "";

            foreach (string parameterName in new[] { "IfcGUID", "IFC GUID", "Ifc GUID" })
            {
                try
                {
                    Parameter p = elem.LookupParameter(parameterName);
                    string value = p == null ? "" : (p.AsString() ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
                catch
                {
                }
            }

            try
            {
                foreach (Parameter p in elem.Parameters)
                {
                    string name = p?.Definition?.Name ?? "";
                    if (name.IndexOf("GUID", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    string value = TryGetParameterStringValue(p);
                    if (!string.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }
            }
            catch
            {
            }

            try
            {
                Element type = doc == null ? null : doc.GetElement(elem.GetTypeId());
                if (type != null && type.Id != elem.Id)
                    return GetElementIfcGuid(null, type);
            }
            catch
            {
            }

            return "";
        }

        private static string TryGetParameterStringValue(Parameter parameter)
        {
            if (parameter == null)
                return "";

            try
            {
                string value = parameter.AsString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch
            {
            }

            try
            {
                string value = parameter.AsValueString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch
            {
            }

            return "";
        }

        private static void GetStairsLevelIds(Stairs stairs, out ElementId baseLevelId, out ElementId topLevelId)
        {
            baseLevelId = null;
            topLevelId = null;
            if (stairs == null)
                return;

            baseLevelId = TryGetElementIdProperty(stairs, "BaseLevelId")
                ?? TryGetElementIdProperty(stairs, "LevelId");
            topLevelId = TryGetElementIdProperty(stairs, "TopLevelId");

            if (baseLevelId == null || baseLevelId == ElementId.InvalidElementId)
                baseLevelId = TryGetElementIdParameterByNames(stairs, "Базовый уровень", "Base Level", "Base Constraint");

            if (topLevelId == null || topLevelId == ElementId.InvalidElementId)
                topLevelId = TryGetElementIdParameterByNames(stairs, "Верхний уровень", "Top Level", "Top Constraint");

            if (baseLevelId == null || baseLevelId == ElementId.InvalidElementId)
                baseLevelId = TryGetElementIdParameterByKeywords(stairs, "баз", "base", "уров", "level");

            if (topLevelId == null || topLevelId == ElementId.InvalidElementId)
                topLevelId = TryGetElementIdParameterByKeywords(stairs, "верх", "top", "уров", "level");
        }

        private static ElementId TryGetElementIdParameterByNames(Element elem, params string[] names)
        {
            if (elem == null || names == null)
                return null;

            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                try
                {
                    Parameter p = elem.LookupParameter(name);
                    ElementId id = TryGetParameterElementId(p);
                    if (id != null && id != ElementId.InvalidElementId)
                        return id;
                }
                catch
                {
                }
            }

            try
            {
                foreach (Parameter p in elem.Parameters)
                {
                    string parameterName = p?.Definition?.Name ?? "";
                    if (!names.Any(x => string.Equals(x, parameterName, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    ElementId id = TryGetParameterElementId(p);
                    if (id != null && id != ElementId.InvalidElementId)
                        return id;
                }
            }
            catch
            {
            }

            return null;
        }

        private static ElementId TryGetElementIdParameterByKeywords(Element elem, params string[] keywords)
        {
            if (elem == null || keywords == null || keywords.Length == 0)
                return null;

            try
            {
                foreach (Parameter p in elem.Parameters)
                {
                    string parameterName = p?.Definition?.Name ?? "";
                    if (string.IsNullOrWhiteSpace(parameterName))
                        continue;

                    string lower = parameterName.ToLowerInvariant();
                    bool hasRussianLevelWord = lower.IndexOf("уров", StringComparison.Ordinal) >= 0;
                    bool hasEnglishLevelWord = lower.IndexOf("level", StringComparison.Ordinal) >= 0;
                    bool hasBaseWord = lower.IndexOf("баз", StringComparison.Ordinal) >= 0 || lower.IndexOf("base", StringComparison.Ordinal) >= 0;
                    bool hasTopWord = lower.IndexOf("верх", StringComparison.Ordinal) >= 0 || lower.IndexOf("top", StringComparison.Ordinal) >= 0;

                    bool wantsBase = keywords.Any(x => string.Equals(x, "баз", StringComparison.OrdinalIgnoreCase) || string.Equals(x, "base", StringComparison.OrdinalIgnoreCase));
                    bool wantsTop = keywords.Any(x => string.Equals(x, "верх", StringComparison.OrdinalIgnoreCase) || string.Equals(x, "top", StringComparison.OrdinalIgnoreCase));

                    if ((wantsBase && hasBaseWord || wantsTop && hasTopWord) && (hasRussianLevelWord || hasEnglishLevelWord))
                    {
                        ElementId id = TryGetParameterElementId(p);
                        if (id != null && id != ElementId.InvalidElementId)
                            return id;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static ElementId TryGetParameterElementId(Parameter parameter)
        {
            if (parameter == null)
                return null;

            try
            {
                ElementId id = parameter.AsElementId();
                return id == ElementId.InvalidElementId ? null : id;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsLevelInSet(ElementId levelId, HashSet<long> levels)
        {
            return levelId != null
                && levelId != ElementId.InvalidElementId
                && levels != null
                && levels.Contains(IDHelper.ElIdValue(levelId));
        }

        private static bool BoundingBoxesIntersectXY(BoundingBoxXYZ a, BoundingBoxXYZ b, double toleranceFt)
        {
            if (a == null || b == null)
                return false;

            double tol = Math.Max(0.0, toleranceFt);
            return a.Min.X <= b.Max.X + tol && a.Max.X >= b.Min.X - tol
                && a.Min.Y <= b.Max.Y + tol && a.Max.Y >= b.Min.Y - tol;
        }

        private static string FormatOptionalElementId(ElementId id)
        {
            return id == null || id == ElementId.InvalidElementId
                ? "нет"
                : IDHelper.ElIdValue(id).ToString();
        }

        private static string FormatYesNo(bool value)
        {
            return value ? "да" : "нет";
        }

        private static void AddElementIds(List<ElementId> result, IEnumerable<ElementId> ids)
        {
            if (result == null || ids == null)
                return;

            foreach (ElementId id in ids)
                AddElementIdIfValid(result, id);
        }

        private static void AddElementIdIfValid(List<ElementId> result, ElementId id)
        {
            if (result == null || id == null || id == ElementId.InvalidElementId)
                return;

            long value = IDHelper.ElIdValue(id);
            if (result.Any(x => IDHelper.ElIdValue(x) == value))
                return;

            result.Add(id);
        }

        private static ElementId TryGetElementIdProperty(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            try
            {
                PropertyInfo property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                return property?.GetValue(target, null) as ElementId;
            }
            catch
            {
                return null;
            }
        }

        private static List<ElementId> TryInvokeElementIdCollection(object target, string methodName, params object[] args)
        {
            var result = new List<ElementId>();
            if (target == null || string.IsNullOrWhiteSpace(methodName))
                return result;

            args = args ?? new object[0];

            try
            {
                MethodInfo method = target.GetType()
                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .FirstOrDefault(x => string.Equals(x.Name, methodName, StringComparison.Ordinal) && x.GetParameters().Length == args.Length);

                if (method == null)
                    return result;

                object value = method.Invoke(target, args);
                if (value == null)
                    return result;

                ElementId single = value as ElementId;
                if (single != null)
                {
                    result.Add(single);
                    return result;
                }

                Element singleElement = value as Element;
                if (singleElement != null)
                {
                    result.Add(singleElement.Id);
                    return result;
                }

                IEnumerable<ElementId> typed = value as IEnumerable<ElementId>;
                if (typed != null)
                    return typed.Where(x => x != null && x != ElementId.InvalidElementId).ToList();

                System.Collections.IEnumerable enumerable = value as System.Collections.IEnumerable;
                if (enumerable == null)
                    return result;

                foreach (object item in enumerable)
                {
                    ElementId id = item as ElementId;
                    if (id != null && id != ElementId.InvalidElementId)
                    {
                        result.Add(id);
                        continue;
                    }

                    Element elem = item as Element;
                    if (elem != null && elem.Id != null && elem.Id != ElementId.InvalidElementId)
                        result.Add(elem.Id);
                }
            }
            catch
            {
                return new List<ElementId>();
            }

            return result;
        }

        private static List<Stairs> GetStairsIntersectingBoundingBox(Document doc, Element elem)
        {
            var result = new List<Stairs>();
            BoundingBoxXYZ bb = SafeGetBoundingBox(elem);
            if (doc == null || bb == null)
                return result;

            try
            {
                foreach (Stairs stairs in new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .OfType<Stairs>())
                {
                    BoundingBoxXYZ stairBb = SafeGetBoundingBox(stairs);
                    if (stairBb == null)
                        continue;

                    if (BoundingBoxesIntersect(bb, stairBb))
                        result.Add(stairs);
                }
            }
            catch
            {
                return new List<Stairs>();
            }

            return result.OrderBy(x => IDHelper.ElIdValue(x.Id)).ToList();
        }

        private static BoundingBoxXYZ SafeGetBoundingBox(Element elem)
        {
            try
            {
                return elem?.get_BoundingBox(null);
            }
            catch
            {
                return null;
            }
        }

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null)
                return false;

            return a.Min.X <= b.Max.X && a.Max.X >= b.Min.X
                && a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y
                && a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z;
        }

        private static long? PickStairElementId(UIApplication uiapp, Document doc)
        {
            try
            {
                var uidoc = uiapp.ActiveUIDocument;
                Reference r = uidoc.Selection.PickObject(ObjectType.Element, new StairsSelectionFilter(doc), "Выберите лестницу или многоэтажную лестницу (Esc — Отмена)");
                return r == null ? (long?)null : IDHelper.ElIdValue(r.ElementId);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        private static void SelectAndShowElement(UIDocument uidoc, long elementId)
        {
            if (uidoc == null || uidoc.Document == null)
                return;

            ElementId id = IDHelper.CreateElementId(elementId);
            if (uidoc.Document.GetElement(id) == null)
                throw new InvalidOperationException($"Элемент ID {elementId} не найден в документе.");

            uidoc.Selection.SetElementIds(new List<ElementId> { id });
            try { uidoc.ShowElements(id); } catch { }
        }

        private static List<ElementId> GetMultistoryStairsIds(MultistoryStairs multistory)
        {
            if (multistory == null)
                return new List<ElementId>();

            try
            {
                ICollection<ElementId> ids = multistory.GetAllStairsIds();
                return ids == null ? new List<ElementId>() : ids.ToList();
            }
            catch
            {
                return new List<ElementId>();
            }
        }

        private static List<ElementId> GetMultistoryConnectedLevelIds(MultistoryStairs multistory)
        {
            return TryInvokeElementIdCollection(multistory, "GetAllConnectedLevels");
        }

        private static string GetElementTypeDisplayName(Document doc, Element elem)
        {
            if (doc == null || elem == null)
                return "";

            try
            {
                ElementId typeId = elem.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                    return "";

                Element type = doc.GetElement(typeId);
                return type?.Name ?? "";
            }
            catch
            {
                return "";
            }
        }


        private static List<EvacuationRoutesWorksetOption> GetEvacuationWorksetOptions(Document doc)
        {
            var result = new List<EvacuationRoutesWorksetOption>();
            if (doc == null || !doc.IsWorkshared)
                return result;

            try
            {
                foreach (Workset ws in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
                {
                    if (ws == null || string.IsNullOrWhiteSpace(ws.Name)) continue;
                    if (ws.Name.IndexOf("ЭВАКУАЦИИ", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    result.Add(new EvacuationRoutesWorksetOption(ws.Id.IntegerValue, ws.Name));
                }
            }
            catch
            {
                return new List<EvacuationRoutesWorksetOption>();
            }

            return result;
        }

        private static string FormatIdsLine(string title, IEnumerable<int> ids)
        {
            var list = (ids ?? Enumerable.Empty<int>())
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            return list.Count == 0
                ? $"{title}: нет"
                : $"{title} (ID): {string.Join(", ", list)}";
        }

        private static bool TooManyIdsForDialog(IEnumerable<int> ids, int countThreshold = 30, int textThreshold = 700)
        {
            var list = (ids ?? Enumerable.Empty<int>()).Distinct().ToList();
            if (list.Count > countThreshold) return true;

            string s = string.Join(", ", list.OrderBy(x => x));
            return s.Length > textThreshold;
        }

        private static string SaveFailuresLogToDesktop(IEnumerable<int> stairIds, IEnumerable<int> runIds, IEnumerable<int> landingIds)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"KPLN_EvacuationRoutes_Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string path = System.IO.Path.Combine(desktop, fileName);

            var sIds = (stairIds ?? Enumerable.Empty<int>()).Distinct().OrderBy(x => x).ToList();
            var rIds = (runIds ?? Enumerable.Empty<int>()).Distinct().OrderBy(x => x).ToList();
            var lIds = (landingIds ?? Enumerable.Empty<int>()).Distinct().OrderBy(x => x).ToList();

            var lines = new List<string>
            {
                "KPLN. Пути эвакуации — лог необработанных элементов",
                $"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                "",
                sIds.Count == 0 ? "Необработанные лестницы: нет" : $"Необработанные лестницы (ID): {string.Join(", ", sIds)}",
                rIds.Count == 0 ? "Необработанные марши: нет"   : $"Необработанные марши (ID): {string.Join(", ", rIds)}",
                lIds.Count == 0 ? "Необработанные площадки: нет": $"Необработанные площадки (ID): {string.Join(", ", lIds)}",
            };

            System.IO.File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
            return path;
        }

        private static void ShowIntersectionReport(UIDocument uidoc, List<RouteIntersectionReportItem> reports, RouteDebugLog debugLog)
        {
            try
            {
                var window = new RouteIntersectionReportWindow(uidoc, reports, debugLog);
                if (uiapp != null && uiapp.MainWindowHandle != IntPtr.Zero)
                    new WindowInteropHelper(window) { Owner = uiapp.MainWindowHandle };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                string reportText = FormatIntersectionReport(reports);
                TaskDialog.Show("Проверить пересечения", $"{reportText}\n\nНе удалось открыть окно выбора:\n{ex.Message}");
            }
        }

        private static string SaveIntersectionReportToDesktop(List<RouteIntersectionReportItem> reports)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"KPLN_EvacuationRoutes_Intersections_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string path = System.IO.Path.Combine(desktop, fileName);

            var lines = new List<string>
            {
                "KPLN. Пути эвакуации — отчёт пересечений",
                $"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                ""
            };

            lines.AddRange(BuildProblemReportLines(null, reports));
            if (lines.Count == 3)
                lines.Add("Пересечений с элементами не найдено.");

            System.IO.File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
            return path;
        }

        private static string FormatIntersectionReport(List<RouteIntersectionReportItem> reports)
        {
            return string.Join("\n", BuildProblemReportLines(null, reports));
        }

        private static IEnumerable<string> FormatIntersectionReportLines(List<RouteIntersectionReportItem> reports)
        {
            foreach (var report in reports ?? new List<RouteIntersectionReportItem>())
            {
                if (report == null || report.Targets == null || report.Targets.Count == 0)
                    continue;

                yield return $"{report.RouteName} (ID: {report.RouteElementId})";

                foreach (var target in report.Targets.OrderBy(x => x.ElementId))
                {
                    string source = string.IsNullOrWhiteSpace(target.SourceName) ? "Host" : target.SourceName;
                    string link = target.LinkInstanceId.HasValue ? $" | LinkInstanceId {target.LinkInstanceId.Value}" : "";
                    string cat = string.IsNullOrWhiteSpace(target.CategoryName) ? "без категории" : target.CategoryName;
                    string name = string.IsNullOrWhiteSpace(target.ElementName) ? "" : $" | {target.ElementName}";
                    yield return $"  {source}{link} | ID {target.ElementId} | {cat}{name}";
                }

                yield return "";
            }
        }

        private static string SaveDebugLogToDesktop(RouteDebugLog debugLog)
        {
            if (debugLog == null || !debugLog.Enabled || debugLog.Lines == null || debugLog.Lines.Count == 0)
                throw new InvalidOperationException("Отчёт о марше доступен только для одиночного запуска с выбранной лестницей.");

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"KPLN_EvacuationRoutes_Debug_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string path = System.IO.Path.Combine(desktop, fileName);
            System.IO.File.WriteAllLines(path, debugLog.Lines, System.Text.Encoding.UTF8);
            return path;
        }

        private static void AddStairDebugHeader(RouteDebugLog debugLog, Stairs stairs)
        {
            if (debugLog == null || !debugLog.Enabled || stairs == null)
                return;

            debugLog.Add("===== ЛЕСТНИЦА =====");
            debugLog.Add($"StairsId={IDHelper.ElIdValue(stairs.Id)}");
            try
            {
                MultistoryStairs parent = GetParentMultistoryStairs(stairs.Document, stairs);
                debugLog.Add($"MultistoryStairsId={(parent == null ? "нет" : IDHelper.ElIdValue(parent.Id).ToString())}");
            }
            catch (Exception ex) { debugLog.Add($"MultistoryStairsId ERROR: {ex.Message}"); }

            try { debugLog.Add($"RunIds={FormatElementIds(stairs.GetStairsRuns())}"); }
            catch (Exception ex) { debugLog.Add($"RunIds ERROR: {ex.Message}"); }

            try { debugLog.Add($"LandingIds={FormatElementIds(stairs.GetStairsLandings())}"); }
            catch (Exception ex) { debugLog.Add($"LandingIds ERROR: {ex.Message}"); }

            BoundingBoxXYZ bb = null;
            try { bb = stairs.get_BoundingBox(null); } catch { }
            debugLog.Add($"BoundingBox={FormatBoundingBox(bb)}");
            debugLog.AddBlank();
        }

        private static void AddRouteBuildTargetDebugHeader(RouteDebugLog debugLog, RouteBuildTarget target)
        {
            if (debugLog == null || !debugLog.Enabled || target == null || target.Stairs == null)
                return;

            debugLog.Add("===== ЦЕЛЬ ОБРАБОТКИ =====");
            debugLog.Add($"DisplayName={target.DisplayName}");
            debugLog.Add($"IsMultistoryPlacement={target.IsMultistoryPlacement}");
            debugLog.Add($"OwnerElementId={target.OwnerElementId}");
            debugLog.Add($"StandardStairsId={target.StandardStairsId}");
            debugLog.Add($"PlacementLevelId={FormatOptionalElementId(target.PlacementLevelId)}");
            debugLog.Add($"VerticalOffset={FormatFtMm(target.VerticalOffsetFt)}");
            AddStairDebugHeader(debugLog, target.Stairs);
        }

        private static string FormatElementIds(ICollection<ElementId> ids)
        {
            if (ids == null || ids.Count == 0)
                return "нет";

            return string.Join(", ", ids.Select(x => IDHelper.ElIdValue(x)).OrderBy(x => x));
        }

        private static string FormatFt(double value)
        {
            return value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string FormatMm(double valueFt)
        {
            return IDHelper.ConvertInternalToMm(valueFt).ToString("0.#", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string FormatFtMm(double valueFt)
        {
            return $"{FormatFt(valueFt)} ft / {FormatMm(valueFt)} mm";
        }

        private static string FormatXyz(XYZ p)
        {
            if (p == null)
                return "null";

            return $"({FormatFt(p.X)}, {FormatFt(p.Y)}, {FormatFt(p.Z)}) ft";
        }

        private static string FormatBoundingBox(BoundingBoxXYZ bb)
        {
            if (bb == null)
                return "null";

            return $"Min={FormatXyz(bb.Min)}; Max={FormatXyz(bb.Max)}";
        }

        private static List<Stairs> GetRouteTargetStairs(Document doc)
        {
            var result = new List<Stairs>();
            var seen = new HashSet<long>();
            if (doc == null)
                return result;

            foreach (Stairs stairs in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .OfType<Stairs>())
            {
                AddStairsIfNew(stairs, result, seen);
            }

            foreach (MultistoryStairs multistory in new FilteredElementCollector(doc)
                .OfClass(typeof(MultistoryStairs))
                .WhereElementIsNotElementType()
                .OfType<MultistoryStairs>())
            {
                AddMultistoryStairsMembers(doc, multistory, result, seen);
            }

            return result.OrderBy(x => IDHelper.ElIdValue(x.Id)).ToList();
        }

        private static List<Stairs> PickSingleStairs(UIApplication uiapp, Document doc, RouteDebugLog debugLog)
        {
            try
            {
                var uidoc = uiapp.ActiveUIDocument;

                Reference r = uidoc.Selection.PickObject(ObjectType.Element, new StairsSelectionFilter(doc), "Выберите лестницу или многоэтажную лестницу (Esc — Отмена)");

                if (r == null) return null;

                Element picked = doc.GetElement(r.ElementId);
                var stairs = GetStairsFromPickedElement(doc, picked);

                debugLog?.Add("===== ВЫБОР =====");
                debugLog?.Add($"PickedElementId={IDHelper.ElIdValue(r.ElementId)}");
                debugLog?.Add($"PickedElementType={(picked == null ? "null" : picked.GetType().Name)}");
                debugLog?.Add($"ExpandedStairs={FormatElementIds(stairs.Select(x => x.Id).ToList())}");
                debugLog?.AddBlank();

                return stairs;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        private static List<Stairs> GetStairsFromPickedElement(Document doc, Element picked)
        {
            var result = new List<Stairs>();
            var seen = new HashSet<long>();
            if (doc == null || picked == null)
                return result;

            MultistoryStairs multistory = picked as MultistoryStairs;
            if (multistory != null)
            {
                AddMultistoryStairsMembers(doc, multistory, result, seen);
                return result.OrderBy(x => IDHelper.ElIdValue(x.Id)).ToList();
            }

            Stairs stairs = picked as Stairs;
            if (stairs == null)
                return result;

            MultistoryStairs parent = GetParentMultistoryStairs(doc, stairs);
            if (parent != null)
                AddMultistoryStairsMembers(doc, parent, result, seen);
            else
                AddStairsIfNew(stairs, result, seen);

            return result.OrderBy(x => IDHelper.ElIdValue(x.Id)).ToList();
        }

        private static MultistoryStairs GetParentMultistoryStairs(Document doc, Stairs stairs)
        {
            if (doc == null || stairs == null)
                return null;

            try
            {
                long stairsId = IDHelper.ElIdValue(stairs.Id);

                foreach (MultistoryStairs multistory in new FilteredElementCollector(doc)
                    .OfClass(typeof(MultistoryStairs))
                    .WhereElementIsNotElementType()
                    .OfType<MultistoryStairs>())
                {
                    List<ElementId> ids = GetMultistoryStandardStairIds(doc, multistory);
                    if (ids.Any(x => IDHelper.ElIdValue(x) == stairsId))
                        return multistory;
                }
            }
            catch
            {
            }

            return null;
        }

        private static void AddMultistoryStairsMembers(Document doc, MultistoryStairs multistory, List<Stairs> result, HashSet<long> seen)
        {
            if (doc == null || multistory == null || result == null || seen == null)
                return;

            foreach (ElementId id in GetMultistoryStandardStairIds(doc, multistory))
            {
                Stairs stairs = doc.GetElement(id) as Stairs;
                AddStairsIfNew(stairs, result, seen);
            }
        }

        private static void AddStairsIfNew(Stairs stairs, List<Stairs> result, HashSet<long> seen)
        {
            if (stairs == null || result == null || seen == null)
                return;

            long id = IDHelper.ElIdValue(stairs.Id);
            if (!seen.Add(id))
                return;

            result.Add(stairs);
        }

        // =========================
        // ЛЕСТНИЦА: МАРШИ + ПЛОЩАДКИ
        // =========================

        private bool TryCreateRouteBodyOnStair(Document doc, RouteBuildTarget target, EvacuationRoutesDialogResult data, List<RouteIntersectionReportItem> intersectionReports, RouteDebugLog debugLog,
            out int createdRuns, out int createdLandings, out List<int> failedRunIds, out List<int> failedLandingIds)
        {
            createdRuns = 0;
            createdLandings = 0;
            failedRunIds = new List<int>();
            failedLandingIds = new List<int>();

            Stairs stairs = target?.Stairs;
            if (doc == null || stairs == null || data == null)
                return false;

            var runIds = stairs.GetStairsRuns();
            var landingIds = stairs.GetStairsLandings();

            bool hasRuns = runIds != null && runIds.Count > 0;
            bool hasLandings = landingIds != null && landingIds.Count > 0;

            if (!hasRuns && !hasLandings) return false;

            debugLog?.Add($"===== ОБРАБОТКА {(target.IsMultistoryPlacement ? "РАЗМЕЩЕНИЯ" : "ЛЕСТНИЦЫ")} {IDHelper.ElIdValue(stairs.Id)} =====");
            debugLog?.Add($"Target={target.DisplayName}; OwnerId={target.OwnerElementId}; PlacementLevel={FormatOptionalElementId(target.PlacementLevelId)}; Offset={FormatFtMm(target.VerticalOffsetFt)}");
            debugLog?.Add($"Runs={FormatElementIds(runIds)}");
            debugLog?.Add($"Landings={FormatElementIds(landingIds)}");
            debugLog?.AddBlank();

#if Debug2023 || Debug2024 || Revit2023 || Revit2024
            double heightFt = UnitUtils.ConvertToInternalUnits(data.HeightMm, UnitTypeId.Millimeters);
            if (heightFt <= 1e-9) return false;
            double epsFt = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
#else
            double heightFt = UnitUtils.ConvertToInternalUnits(data.HeightMm, DisplayUnitType.DUT_MILLIMETERS);
            if (heightFt <= 1e-9) return false;
            double epsFt = UnitUtils.ConvertToInternalUnits(1.0, DisplayUnitType.DUT_MILLIMETERS);
#endif
            var runInfos = new List<RunRouteBodyInfo>();

            // МАРШИ
            if (hasRuns)
            {
                foreach (ElementId runId in runIds)
                {
                    StairsRun run = doc.GetElement(runId) as StairsRun;
                    if (run == null)
                    {
                        failedRunIds.Add(IDHelper.ElIdInt(runId));
                        continue;
                    }

                    RunRouteBodyInfo info;
                    debugLog?.Add($"--- МАРШ {IDHelper.ElIdValue(run.Id)} ---");
                    bool okRun = TryCreateRouteBodyOnRun(doc, target, run, data, heightFt, epsFt, intersectionReports, debugLog, out info);
                    debugLog?.Add($"RunResult={okRun}");
                    debugLog?.AddBlank();

                    if (okRun)
                    {
                        createdRuns++;
                        if (info != null) runInfos.Add(info);
                    }
                    else
                    {
                        failedRunIds.Add(IDHelper.ElIdInt(runId));
                    }
                }
            }

            // ПЛОЩАДКИ
            if (hasLandings)
            {
                var runs = new List<StairsRun>();
                if (hasRuns)
                {
                    foreach (var rid in runIds)
                    {
                        var r = doc.GetElement(rid) as StairsRun;
                        if (r != null) runs.Add(r);
                    }
                }

                foreach (ElementId landingId in landingIds)
                {
                    StairsLanding landing = doc.GetElement(landingId) as StairsLanding;
                    if (landing == null)
                    {
                        failedLandingIds.Add(IDHelper.ElIdInt(landingId));
                        continue;
                    }

                    bool okLanding = TryCreateRouteBodyOnLanding(doc, target, landing, runs, runInfos, data, heightFt, intersectionReports, debugLog);
                    debugLog?.Add($"LandingResult landingId={IDHelper.ElIdValue(landing.Id)} ok={okLanding}");
                    debugLog?.AddBlank();
                    if (okLanding) createdLandings++;
                    else failedLandingIds.Add(IDHelper.ElIdInt(landingId));
                }
            }

            return (createdRuns + createdLandings) > 0;
        }

        private static DirectShape FindExistingRouteShape(Document doc, string appId, string appDataId)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(DirectShape)).Cast<DirectShape>().
                FirstOrDefault(ds => string.Equals(ds.ApplicationId, appId, StringComparison.Ordinal) &&
                    string.Equals(ds.ApplicationDataId, appDataId, StringComparison.Ordinal));
        }

        // Создаёт или обновляет DirectShape с заданным appId/appDataId.
        private static DirectShape UpsertRouteShape(Document doc, ElementId categoryId, string appId, string appDataId, string name, Solid solid, EvacuationRoutesDialogResult data)
        {
            if (solid == null || solid.Volume < 1e-9)
                return null;

            DirectShape ds = FindExistingRouteShape(doc, appId, appDataId);

            if (ds == null)
            {
                ds = DirectShape.CreateElement(doc, categoryId);
                ds.ApplicationId = appId;
                ds.ApplicationDataId = appDataId;
            }

            ds.Name = name;
            ds.SetShape(new List<GeometryObject> { solid });
            TrySetRouteShapeWorkset(ds, data);
            return ds;
        }

        private static void TrySetRouteShapeWorkset(DirectShape ds, EvacuationRoutesDialogResult data)
        {
            if (ds == null || data == null || !data.AddToEvacuationWorkset || !data.EvacuationWorksetId.HasValue)
                return;

            Parameter p = ds.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            if (p == null || p.IsReadOnly)
                return;

            try
            {
                p.Set(data.EvacuationWorksetId.Value);
            }
            catch
            {
            }
        }

        private static HashSet<long> GetStairAndComponentIds(Stairs stairs)
        {
            var ids = new HashSet<long>();
            if (stairs == null)
                return ids;

            AddElementAndDependentsToExclude(stairs, ids, 0);

            try
            {
                var runIds = stairs.GetStairsRuns();
                if (runIds != null)
                {
                    foreach (ElementId id in runIds)
                    {
                        Element elem = stairs.Document?.GetElement(id);
                        AddElementAndDependentsToExclude(elem, ids, 0);
                    }
                }
            }
            catch
            {
            }

            try
            {
                var landingIds = stairs.GetStairsLandings();
                if (landingIds != null)
                {
                    foreach (ElementId id in landingIds)
                    {
                        Element elem = stairs.Document?.GetElement(id);
                        AddElementAndDependentsToExclude(elem, ids, 0);
                    }
                }
            }
            catch
            {
            }

            AddAssociatedRailingIds(stairs, ids);

            return ids;
        }

        private static void AddElementAndDependentsToExclude(Element elem, HashSet<long> ids, int depth)
        {
            if (elem == null || ids == null)
                return;

            long id = IDHelper.ElIdValue(elem.Id);
            if (id > 0)
                ids.Add(id);

            if (depth >= 4)
                return;

            ICollection<ElementId> dependentIds = null;
            try
            {
                dependentIds = elem.GetDependentElements(null);
            }
            catch
            {
            }

            if (dependentIds == null || dependentIds.Count == 0)
                return;

            Document doc = elem.Document;
            foreach (ElementId dependentId in dependentIds)
            {
                long dependentLongId = IDHelper.ElIdValue(dependentId);
                if (dependentLongId <= 0 || ids.Contains(dependentLongId))
                    continue;

                Element dependent = null;
                try { dependent = doc?.GetElement(dependentId); } catch { }
                if (dependent == null)
                    ids.Add(dependentLongId);
                else
                    AddElementAndDependentsToExclude(dependent, ids, depth + 1);
            }
        }

        private static void AddAssociatedRailingIds(Stairs stairs, HashSet<long> ids)
        {
            if (stairs == null || ids == null)
                return;

            try
            {
                MethodInfo method = typeof(Railing)
                    .GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .FirstOrDefault(x =>
                    {
                        if (!string.Equals(x.Name, "GetAssociatedRailings", StringComparison.Ordinal))
                            return false;

                        ParameterInfo[] parameters = x.GetParameters();
                        return parameters.Length == 2
                            && parameters[0].ParameterType == typeof(Document)
                            && parameters[1].ParameterType == typeof(ElementId);
                    });

                if (method == null)
                    return;

                object raw = method.Invoke(null, new object[] { stairs.Document, stairs.Id });
                IEnumerable<ElementId> railingIds = raw as IEnumerable<ElementId>;
                if (railingIds == null)
                    return;

                foreach (ElementId railingId in railingIds)
                {
                    Element railing = null;
                    try { railing = stairs.Document?.GetElement(railingId); } catch { }
                    if (railing == null)
                        ids.Add(IDHelper.ElIdValue(railingId));
                    else
                        AddElementAndDependentsToExclude(railing, ids, 0);
                }
            }
            catch
            {
            }
        }

        private static void AddRouteIntersectionReport(Document doc, Solid routeSolid, DirectShape routeShape, string routeName, HashSet<long> excludedIds, List<RouteIntersectionReportItem> reports, RouteDebugLog debugLog, RouteBuildTarget target, ElementId componentId, string componentKind)
        {
            if (doc == null || routeSolid == null || routeSolid.Volume < 1e-9 || reports == null)
                return;

            var targets = new List<RouteIntersectionTarget>();
            var seen = new HashSet<string>();

            int hostRawCount = 0;
            int hostReportCount = 0;
            debugLog?.Add($"[INTERSECTION] {routeName}: current stair excluded ids={(excludedIds == null ? 0 : excludedIds.Count)}");
            try
            {
                var intersected = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementIntersectsSolidFilter(routeSolid))
                    .ToElements();

                hostRawCount = intersected.Count;

                foreach (Element elem in intersected)
                {
                    if (!IsReportableIntersectionElement(elem, excludedIds))
                        continue;

                    if (!HasMeaningfulSolidIntersection(routeSolid, elem))
                        continue;

                    long id = IDHelper.ElIdValue(elem.Id);
                    if (!seen.Add("host:" + id))
                        continue;

                    targets.Add(new RouteIntersectionTarget
                    {
                        SourceName = "Host",
                        ElementId = id,
                        CategoryName = GetElementCategoryName(elem),
                        ElementName = GetElementDisplayName(elem)
                    });

                    hostReportCount++;
                }
            }
            catch (Exception ex)
            {
                debugLog?.Add($"[INTERSECTION] {routeName}: host check ERROR: {ex.Message}");
            }

            debugLog?.Add($"[INTERSECTION] {routeName}: host raw={hostRawCount}; reportable={hostReportCount}");

            debugLog?.Add($"[INTERSECTION] {routeName}: linked documents are skipped");

            if (targets.Count == 0)
                return;

            reports.Add(new RouteIntersectionReportItem
            {
                RouteName = string.IsNullOrWhiteSpace(routeName) ? "Путь эвакуации" : routeName,
                RouteElementId = routeShape == null ? -1 : IDHelper.ElIdValue(routeShape.Id),
                OwnerElementId = target == null ? 0 : target.OwnerElementId,
                ComponentElementId = IDHelper.ElIdValue(componentId),
                ComponentKind = componentKind ?? "",
                PlacementLevelId = target == null ? ElementId.InvalidElementId : target.PlacementLevelId,
                Targets = targets.OrderBy(x => x.ElementId).ToList()
            });
        }

        private static bool IsReportableIntersectionElement(Element elem, HashSet<long> excludedIds)
        {
            if (elem == null)
                return false;

            if (elem is RevitLinkInstance)
                return false;

            long id = IDHelper.ElIdValue(elem.Id);
            if (excludedIds != null && excludedIds.Contains(id))
                return false;

            if (IsOwnRouteShape(elem))
                return false;

            Category cat = elem.Category;
            if (cat == null || cat.CategoryType != CategoryType.Model)
                return false;

            return true;
        }

        private static bool HasMeaningfulSolidIntersection(Solid routeSolid, Element elem)
        {
            if (routeSolid == null || routeSolid.Volume < 1e-9 || elem == null)
                return false;

            var solids = new List<Solid>();
            AddElementSolids(elem, solids);
            if (solids.Count == 0)
                return false;

            double minIntersectionVolume = GetMinIntersectionVolumeFt3();

            foreach (Solid elemSolid in solids)
            {
                if (elemSolid == null || elemSolid.Volume < 1e-9)
                    continue;

                try
                {
                    Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(routeSolid, elemSolid, BooleanOperationsType.Intersect);
                    if (intersection != null && intersection.Volume > minIntersectionVolume && HasMinimumIntersectionThickness(intersection))
                        return true;
                }
                catch
                {
                    // Если boolean не смог посчитать пересечение, не считаем касание доказанным пересечением.
                }
            }

            return false;
        }

        private static double GetMinIntersectionVolumeFt3()
        {
            double oneMmFt = MmToInternal(1.0);
            return oneMmFt * oneMmFt * oneMmFt;
        }

        private static bool HasMinimumIntersectionThickness(Solid intersection)
        {
            if (intersection == null || intersection.Volume <= 1e-12)
                return false;

            BoundingBoxXYZ box;
            try
            {
                box = intersection.GetBoundingBox();
            }
            catch
            {
                return true;
            }

            if (box == null)
                return true;

            double dx = Math.Abs(box.Max.X - box.Min.X);
            double dy = Math.Abs(box.Max.Y - box.Min.Y);
            double dz = Math.Abs(box.Max.Z - box.Min.Z);
            double maxFaceArea = Math.Max(dx * dy, Math.Max(dx * dz, dy * dz));
            if (maxFaceArea <= 1e-12)
                return false;

            double effectiveThicknessFt = intersection.Volume / maxFaceArea;
            return effectiveThicknessFt >= GetMinIntersectionThicknessFt();
        }

        private static double GetMinIntersectionThicknessFt()
        {
            return MmToInternal(5.0);
        }

        private static void AddLinkedIntersectionTargets(Document hostDoc, Solid routeSolid, string routeName, List<RouteIntersectionTarget> targets, HashSet<string> seen, RouteDebugLog debugLog)
        {
            if (hostDoc == null || routeSolid == null || targets == null || seen == null)
                return;

            List<RevitLinkInstance> links;
            try
            {
                links = new FilteredElementCollector(hostDoc)
                    .OfClass(typeof(RevitLinkInstance))
                    .WhereElementIsNotElementType()
                    .Cast<RevitLinkInstance>()
                    .ToList();
            }
            catch (Exception ex)
            {
                debugLog?.Add($"[INTERSECTION] {routeName}: link collector ERROR: {ex.Message}");
                return;
            }

            debugLog?.Add($"[INTERSECTION] {routeName}: link instances={links.Count}");

            foreach (RevitLinkInstance link in links)
            {
                if (link == null)
                    continue;

                Document linkDoc = null;
                try { linkDoc = link.GetLinkDocument(); } catch { }
                if (linkDoc == null)
                {
                    debugLog?.Add($"[INTERSECTION] {routeName}: link {IDHelper.ElIdValue(link.Id)} '{GetElementDisplayName(link)}' has no loaded document");
                    continue;
                }

                Solid linkSolid;
                try
                {
                    Transform toLink = link.GetTransform().Inverse;
                    linkSolid = SolidUtils.CreateTransformed(routeSolid, toLink);
                }
                catch (Exception ex)
                {
                    debugLog?.Add($"[INTERSECTION] {routeName}: link {IDHelper.ElIdValue(link.Id)} transform ERROR: {ex.Message}");
                    continue;
                }

                int rawCount = 0;
                int reportCount = 0;

                try
                {
                    var intersected = new FilteredElementCollector(linkDoc)
                        .WhereElementIsNotElementType()
                        .WherePasses(new ElementIntersectsSolidFilter(linkSolid))
                        .ToElements();

                    rawCount = intersected.Count;

                    foreach (Element elem in intersected)
                    {
                        if (!IsReportableIntersectionElement(elem, null))
                            continue;

                        if (!HasMeaningfulSolidIntersection(linkSolid, elem))
                            continue;

                        long elemId = IDHelper.ElIdValue(elem.Id);
                        long linkId = IDHelper.ElIdValue(link.Id);
                        string key = "link:" + linkId + ":" + elemId;
                        if (!seen.Add(key))
                            continue;

                        targets.Add(new RouteIntersectionTarget
                        {
                            SourceName = string.IsNullOrWhiteSpace(linkDoc.Title) ? GetElementDisplayName(link) : linkDoc.Title,
                            LinkInstanceId = linkId,
                            ElementId = elemId,
                            CategoryName = GetElementCategoryName(elem),
                            ElementName = GetElementDisplayName(elem)
                        });

                        reportCount++;
                    }
                }
                catch (Exception ex)
                {
                    debugLog?.Add($"[INTERSECTION] {routeName}: link {IDHelper.ElIdValue(link.Id)} '{linkDoc.Title}' check ERROR: {ex.Message}");
                    continue;
                }

                debugLog?.Add($"[INTERSECTION] {routeName}: link {IDHelper.ElIdValue(link.Id)} '{linkDoc.Title}' raw={rawCount}; reportable={reportCount}");
            }
        }

        private static string GetElementCategoryName(Element elem)
        {
            try
            {
                return elem?.Category?.Name ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string GetElementDisplayName(Element elem)
        {
            try
            {
                return elem?.Name ?? "";
            }
            catch
            {
                return "";
            }
        }

        // =========================
        // МАРШ
        // =========================
        private bool TryCreateRouteBodyOnRun(Document doc, RouteBuildTarget target, StairsRun run, EvacuationRoutesDialogResult data, double heightFt, double epsFt, List<RouteIntersectionReportItem> intersectionReports, RouteDebugLog debugLog, out RunRouteBodyInfo runInfo)
        {
            runInfo = null;
            Stairs stairs = target?.Stairs;
            if (stairs == null)
                return false;

            CurveLoop path;
            try { path = run.GetStairsPath(); }
            catch { return false; }

            if (path == null) return false;

            var curves = path.ToList();
            if (curves.Count == 0) return false;

            XYZ p0 = curves.First().GetEndPoint(0);
            XYZ p1 = curves.Last().GetEndPoint(1);

            XYZ bottomCenter, topCenter;
            if (p0.Z <= p1.Z) { bottomCenter = p0; topCenter = p1; }
            else { bottomCenter = p1; topCenter = p0; }

            debugLog?.Add($"PathCurves={curves.Count}");
            debugLog?.Add($"PathP0={FormatXyz(p0)}");
            debugLog?.Add($"PathP1={FormatXyz(p1)}");
            debugLog?.Add($"BottomCenter={FormatXyz(bottomCenter)}");
            debugLog?.Add($"TopCenter={FormatXyz(topCenter)}");

            XYZ xP = new XYZ(topCenter.X - bottomCenter.X, topCenter.Y - bottomCenter.Y, 0.0);
            double lenPlan = xP.GetLength();
            if (lenPlan < 1e-9) return false;
            xP = xP.Normalize();

            XYZ yP = XYZ.BasisZ.CrossProduct(xP);
            if (yP.GetLength() < 1e-9) yP = XYZ.BasisY;
            yP = yP.Normalize();

            double widthFt;
            XYZ routeBottomCenter = bottomCenter;
            XYZ routeTopCenter = topCenter;

            if (data.UseRunWidth)
            {
                RunClearWidthInfo clearWidth = GetRunClearWidthInfo(doc, stairs, run, bottomCenter, topCenter, xP, yP, lenPlan);
                widthFt = clearWidth.WidthFt;
                routeBottomCenter = bottomCenter + yP * clearWidth.CenterOffsetFt;
                routeTopCenter = topCenter + yP * clearWidth.CenterOffsetFt;
                debugLog?.Add($"WidthMode=RunClearWidth; Width={FormatFtMm(widthFt)}; CenterOffset={FormatFtMm(clearWidth.CenterOffsetFt)}");
            }
            else
            {
                widthFt = MmToInternal(data.WidthMm);
                debugLog?.Add($"WidthMode=Manual; Width={FormatFtMm(widthFt)}");
            }

            if (widthFt <= 1e-9) return false;

            debugLog?.Add($"LenPlan={FormatFtMm(lenPlan)}");
            debugLog?.Add($"XDir={FormatXyz(xP)}");
            debugLog?.Add($"YDir={FormatXyz(yP)}");
            debugLog?.Add($"RouteBottomCenter={FormatXyz(routeBottomCenter)}");
            debugLog?.Add($"RouteTopCenter={FormatXyz(routeTopCenter)}");

            XYZ halfW = yP * (widthFt / 2.0);

            Plane undersidePlane;
            if (!TryGetBestUndersidePlane(run, out undersidePlane))
                return false;

            RunTopSearchContext topSearch = BuildRunTopSearchContext(doc, stairs, run);
            double baseMidGapFt = GetMidGapFt(topSearch, undersidePlane, routeBottomCenter, xP, lenPlan, widthFt, yP, includeFinish: false);
            if (baseMidGapFt <= 1e-9) baseMidGapFt = 0.0;

            double finishMidGapFt = GetMidGapFt(topSearch, undersidePlane, routeBottomCenter, xP, lenPlan, widthFt, yP, includeFinish: true);
            if (finishMidGapFt < baseMidGapFt) finishMidGapFt = baseMidGapFt;

            double baseLiftFt = baseMidGapFt + epsFt;
            double liftFt = finishMidGapFt + epsFt;

            debugLog?.Add($"BaseMidGap={FormatFtMm(baseMidGapFt)}");
            debugLog?.Add($"FinishMidGap={FormatFtMm(finishMidGapFt)}");
            debugLog?.Add($"BaseLift={FormatFtMm(baseLiftFt)}");
            debugLog?.Add($"Lift={FormatFtMm(liftFt)}");

            XYZ SL_xy = new XYZ(routeBottomCenter.X, routeBottomCenter.Y, 0) - halfW;
            XYZ SR_xy = new XYZ(routeBottomCenter.X, routeBottomCenter.Y, 0) + halfW;
            XYZ EL_xy = new XYZ(routeTopCenter.X, routeTopCenter.Y, 0) - halfW;
            XYZ ER_xy = new XYZ(routeTopCenter.X, routeTopCenter.Y, 0) + halfW;

            double zSL = GetPlaneZAtXY(undersidePlane, SL_xy.X, SL_xy.Y);
            double zSR = GetPlaneZAtXY(undersidePlane, SR_xy.X, SR_xy.Y);
            double zEL = GetPlaneZAtXY(undersidePlane, EL_xy.X, EL_xy.Y);
            double zER = GetPlaneZAtXY(undersidePlane, ER_xy.X, ER_xy.Y);

            XYZ SL = new XYZ(SL_xy.X, SL_xy.Y, zSL + liftFt);
            XYZ SR = new XYZ(SR_xy.X, SR_xy.Y, zSR + liftFt);
            XYZ ER = new XYZ(ER_xy.X, ER_xy.Y, zER + liftFt);
            XYZ EL = new XYZ(EL_xy.X, EL_xy.Y, zEL + liftFt);

            XYZ SL_base = new XYZ(SL_xy.X, SL_xy.Y, zSL + baseLiftFt);
            XYZ SR_base = new XYZ(SR_xy.X, SR_xy.Y, zSR + baseLiftFt);
            XYZ ER_base = new XYZ(ER_xy.X, ER_xy.Y, zER + baseLiftFt);
            XYZ EL_base = new XYZ(EL_xy.X, EL_xy.Y, zEL + baseLiftFt);

            debugLog?.Add($"RunBottomFace finish SL={FormatXyz(SL)}; SR={FormatXyz(SR)}; EL={FormatXyz(EL)}; ER={FormatXyz(ER)}");
            debugLog?.Add($"RunBottomFace base SL={FormatXyz(SL_base)}; SR={FormatXyz(SR_base)}; EL={FormatXyz(EL_base)}; ER={FormatXyz(ER_base)}");

            XYZ up = XYZ.BasisZ * heightFt;
            XYZ SLt = SL + up;
            XYZ SRt = SR + up;
            XYZ ERt = ER + up;
            XYZ ELt = EL + up;

            Solid solid = BuildPrismFrom8Points(SL, SR, ER, EL, SLt, SRt, ERt, ELt);
            if (solid == null || solid.Volume < 1e-9)
                return false;

            solid = TransformSolidForTarget(solid, target);
            if (solid == null || solid.Volume < 1e-9)
                return false;

            runInfo = new RunRouteBodyInfo
            {
                RunId = run.Id,
                StairsId = IDHelper.ElIdInt(stairs.Id),
                WidthFt = widthFt,
                HeightFt = heightFt,
                XDirPlan = xP,
                YDirPlan = yP,

                BottomEnd = new EndFace
                {
                    BL = SL_base,
                    BR = SR_base,
                    TR = SR_base + up,
                    TL = SL_base + up
                },

                TopEnd = new EndFace
                {
                    BL = EL_base,
                    BR = ER_base,
                    TR = ER_base + up,
                    TL = EL_base + up
                }
            };

            ApplyVerticalOffset(runInfo, target?.VerticalOffsetFt ?? 0.0);

            // СОЗДАТЬ ИЛИ ОБНОВИТЬ
            string routeName = CreateRouteName(target, run.Id, isLanding: false);
            string appDataId = CreateRouteAppDataId(target, run.Id);
            DirectShape routeShape = UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", appDataId, routeName, solid, data);
            AddRouteIntersectionReport(doc, solid, routeShape, routeName, GetStairAndComponentIds(stairs), intersectionReports, debugLog, target, run.Id, "Марш");
            return true;
        }

        // =========================
        // ПЛОЩАДКА: коробка по экстремумам углов двух маршей
        // =========================
        private bool TryCreateRouteBodyOnLanding(Document doc, RouteBuildTarget target, StairsLanding landing, List<StairsRun> runsInSameStair, List<RunRouteBodyInfo> runInfos, EvacuationRoutesDialogResult data, double heightFt, List<RouteIntersectionReportItem> intersectionReports, RouteDebugLog debugLog)
        {
            Stairs stairs = target?.Stairs;
            if (stairs == null)
                return false;

            debugLog?.Add($"--- ПЛОЩАДКА {IDHelper.ElIdValue(landing.Id)} ---");
            debugLog?.Add($"RunInfosCount={(runInfos == null ? 0 : runInfos.Count)}");

            if (runInfos == null || runInfos.Count < 2)
                return false;

            BoundingBoxXYZ bbL = landing.get_BoundingBox(null);
            if (bbL == null) return false;

            XYZ landingCenter = new XYZ(
                (bbL.Min.X + bbL.Max.X) * 0.5,
                (bbL.Min.Y + bbL.Max.Y) * 0.5,
                (bbL.Min.Z + bbL.Max.Z) * 0.5);

            debugLog?.Add($"LandingBoundingBox={FormatBoundingBox(bbL)}");
            debugLog?.Add($"LandingCenter={FormatXyz(landingCenter)}");

            // Для каждого марша выбираем ближайший к площадке торец (Top/Bottom)
            var candidates = new List<(RunRouteBodyInfo run, EndFace face, XYZ faceCenter, double dist, string endName)>();
            foreach (var ri in runInfos)
            {
                double dTop = new XYZ(ri.TopEnd.Center.X - landingCenter.X, ri.TopEnd.Center.Y - landingCenter.Y, 0).GetLength();
                double dBot = new XYZ(ri.BottomEnd.Center.X - landingCenter.X, ri.BottomEnd.Center.Y - landingCenter.Y, 0).GetLength();

                bool useTop = dTop <= dBot;
                EndFace f = useTop ? ri.TopEnd : ri.BottomEnd;

                debugLog?.Add($"Candidate run={IDHelper.ElIdValue(ri.RunId)} dTop={FormatFtMm(dTop)} dBottom={FormatFtMm(dBot)} selected={(useTop ? "Top" : "Bottom")} selectedCenter={FormatXyz(f.Center)} selectedMinZ={FormatFtMm(f.MinZBottom)} width={FormatFtMm(ri.WidthFt)}");

                candidates.Add((ri, f, f.Center, useTop ? dTop : dBot, useTop ? "Top" : "Bottom"));
            }

            if (candidates.Count < 2)
                return false;

            // Берём 2 ближайших к площадке марша/торца
            var two = candidates.OrderBy(x => x.dist).Take(2).ToList();
            var A = two[0];
            var B = two[1];

            debugLog?.Add($"SelectedA run={IDHelper.ElIdValue(A.run.RunId)} end={A.endName} dist={FormatFtMm(A.dist)} center={FormatXyz(A.faceCenter)}");
            debugLog?.Add($"SelectedB run={IDHelper.ElIdValue(B.run.RunId)} end={B.endName} dist={FormatFtMm(B.dist)} center={FormatXyz(B.faceCenter)}");

            // xDir - направление длины блока на площадке
            XYZ xDir = new XYZ(B.faceCenter.X - A.faceCenter.X, B.faceCenter.Y - A.faceCenter.Y, 0);
            if (xDir.GetLength() < 1e-9)
                xDir = new XYZ(A.run.XDirPlan.X, A.run.XDirPlan.Y, 0);
            if (xDir.GetLength() < 1e-9)
                return false;
            xDir = xDir.Normalize();

            debugLog?.Add($"LandingXDir={FormatXyz(xDir)}");

            // yDir - направление глубины блока поперёк марша
            XYZ yBase = new XYZ(A.run.YDirPlan.X, A.run.YDirPlan.Y, 0);
            if (yBase.GetLength() < 1e-9) yBase = new XYZ(B.run.YDirPlan.X, B.run.YDirPlan.Y, 0);
            if (yBase.GetLength() < 1e-9) yBase = XYZ.BasisZ.CrossProduct(xDir);

            debugLog?.Add($"LandingYBaseBeforeOrtho={FormatXyz(yBase)}");

            yBase = yBase.Normalize();
            XYZ yDir = yBase - xDir * (yBase.DotProduct(xDir));
            if (yDir.GetLength() < 1e-9)
                yDir = XYZ.BasisZ.CrossProduct(xDir);

            yDir = new XYZ(yDir.X, yDir.Y, 0);
            if (yDir.GetLength() < 1e-9) return false;
            yDir = yDir.Normalize();

            debugLog?.Add($"LandingYDir={FormatXyz(yDir)}");

            // X-границы: min/max по углам двух блоков маршей
            double minX = double.PositiveInfinity;
            double maxX = double.NegativeInfinity;

            foreach (var p in A.run.GetAll8Corners().Concat(B.run.GetAll8Corners()))
            {
                double tx = new XYZ(p.X, p.Y, 0).DotProduct(xDir);
                if (tx < minX) minX = tx;
                if (tx > maxX) maxX = tx;
            }

            if (maxX - minX <= 1e-9)
                return false;

            debugLog?.Add($"LandingXSpan minX={FormatFt(minX)} maxX={FormatFt(maxX)} width={FormatFtMm(maxX - minX)}");

            // SPAN по Y из углов блоков
            double spanMinY = double.PositiveInfinity;
            double spanMaxY = double.NegativeInfinity;

            foreach (var p in A.run.GetAll8Corners().Concat(B.run.GetAll8Corners()))
            {
                double ty = new XYZ(p.X, p.Y, 0).DotProduct(yDir);
                if (ty < spanMinY) spanMinY = ty;
                if (ty > spanMaxY) spanMaxY = ty;
            }

            if (spanMaxY - spanMinY <= 1e-9)
                return false;

            debugLog?.Add($"LandingYSpanFromRuns spanMinY={FormatFt(spanMinY)} spanMaxY={FormatFt(spanMaxY)} width={FormatFtMm(spanMaxY - spanMinY)}");


#if Debug2023 || Debug2024 || Revit2023 || Revit2024
            double depthFt = data.UseRunWidth ? Math.Max(A.run.WidthFt, B.run.WidthFt) : UnitUtils.ConvertToInternalUnits(data.WidthMm, UnitTypeId.Millimeters);

            if (depthFt <= 1e-9)
                return false;

            double landingCY = new XYZ(landingCenter.X, landingCenter.Y, 0).DotProduct(yDir);
            double tol = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters);
#else
            double depthFt = data.UseRunWidth ? Math.Max(A.run.WidthFt, B.run.WidthFt) : UnitUtils.ConvertToInternalUnits(data.WidthMm, DisplayUnitType.DUT_MILLIMETERS);

            if (depthFt <= 1e-9)
                return false;

            double landingCY = new XYZ(landingCenter.X, landingCenter.Y, 0).DotProduct(yDir);
            double tol = UnitUtils.ConvertToInternalUnits(2.0, DisplayUnitType.DUT_MILLIMETERS);
#endif

            double minY, maxY;
            string yDecision;

            if (landingCY > spanMaxY + tol)
            {
                // Площадка по + стороне: начинаем от грани spanMaxY и уходим к площадке на depth
                minY = spanMaxY;
                maxY = spanMaxY + depthFt;
                yDecision = "landingCY > spanMaxY + tol";
            }
            else if (landingCY < spanMinY - tol)
            {
                // Площадка по - стороне: начинаем от грани spanMinY и уходим к площадке на depth
                maxY = spanMinY;
                minY = spanMinY - depthFt;
                yDecision = "landingCY < spanMinY - tol";
            }
            else
            {
                // Площадка между гранями по yDir: выбираем ближайшую грань к центру площадки
                double distToMin = Math.Abs(landingCY - spanMinY);
                double distToMax = Math.Abs(spanMaxY - landingCY);

                if (distToMax <= distToMin)
                {
                    minY = spanMaxY;
                    maxY = spanMaxY + depthFt;
                    yDecision = $"inside span; chose spanMaxY because distToMax={FormatFtMm(distToMax)} <= distToMin={FormatFtMm(distToMin)}";
                }
                else
                {
                    maxY = spanMinY;
                    minY = spanMinY - depthFt;
                    yDecision = $"inside span; chose spanMinY because distToMin={FormatFtMm(distToMin)} < distToMax={FormatFtMm(distToMax)}";
                }
            }

            debugLog?.Add($"LandingDepth={FormatFtMm(depthFt)} landingCY={FormatFt(landingCY)} tol={FormatFtMm(tol)}");
            debugLog?.Add($"LandingYDecision={yDecision}");
            debugLog?.Add($"LandingYResult minY={FormatFt(minY)} maxY={FormatFt(maxY)} depth={FormatFtMm(maxY - minY)}");

            // Z-низ/высота: старт от нижнего блока, высота = height блока
            double baseZ = Math.Min(A.face.MinZBottom, B.face.MinZBottom);
            double h = heightFt;
            if (h <= 1e-9) return false;

            debugLog?.Add($"LandingBaseZ={FormatFtMm(baseZ)} height={FormatFtMm(h)}");

            XYZ P1 = xDir * minX + yDir * minY + XYZ.BasisZ * baseZ;
            XYZ P2 = xDir * maxX + yDir * minY + XYZ.BasisZ * baseZ;
            XYZ P3 = xDir * maxX + yDir * maxY + XYZ.BasisZ * baseZ;
            XYZ P4 = xDir * minX + yDir * maxY + XYZ.BasisZ * baseZ;

            debugLog?.Add($"LandingPoints P1={FormatXyz(P1)} P2={FormatXyz(P2)} P3={FormatXyz(P3)} P4={FormatXyz(P4)}");

            XYZ up = XYZ.BasisZ * h;

            Solid solid = BuildPrismFrom8Points(P1, P2, P3, P4, P1 + up, P2 + up, P3 + up, P4 + up);
            if (solid == null || solid.Volume < 1e-9)
                return false;

            // СОЗДАТЬ ИЛИ ОБНОВИТЬ
            string routeName = CreateRouteName(target, landing.Id, isLanding: true);
            string appDataId = CreateRouteAppDataId(target, landing.Id);
            DirectShape routeShape = UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", appDataId, routeName, solid, data);
            AddRouteIntersectionReport(doc, solid, routeShape, routeName, GetStairAndComponentIds(stairs), intersectionReports, debugLog, target, landing.Id, "Площадка");
            return true;
        }

        private static Solid TransformSolidForTarget(Solid solid, RouteBuildTarget target)
        {
            if (solid == null || target == null || Math.Abs(target.VerticalOffsetFt) < 1e-9)
                return solid;

            try
            {
                Transform transform = Transform.CreateTranslation(new XYZ(0, 0, target.VerticalOffsetFt));
                return SolidUtils.CreateTransformed(solid, transform);
            }
            catch
            {
                return solid;
            }
        }

        private static void ApplyVerticalOffset(RunRouteBodyInfo info, double offsetFt)
        {
            if (info == null || Math.Abs(offsetFt) < 1e-9)
                return;

            XYZ offset = new XYZ(0, 0, offsetFt);
            info.BottomEnd = OffsetEndFace(info.BottomEnd, offset);
            info.TopEnd = OffsetEndFace(info.TopEnd, offset);
        }

        private static EndFace OffsetEndFace(EndFace face, XYZ offset)
        {
            return new EndFace
            {
                BL = face.BL + offset,
                BR = face.BR + offset,
                TR = face.TR + offset,
                TL = face.TL + offset
            };
        }

        private static string CreateRouteAppDataId(RouteBuildTarget target, ElementId componentId)
        {
            string component = IDHelper.ElIdValue(componentId).ToString();
            if (target == null || !target.IsMultistoryPlacement || Math.Abs(target.VerticalOffsetFt) < 1e-9)
                return component;

            return $"{target.ShapeKeyPrefix}_{component}";
        }

        private static string CreateRouteName(RouteBuildTarget target, ElementId componentId, bool isLanding)
        {
            long stairId = target?.Stairs == null ? 0 : IDHelper.ElIdValue(target.Stairs.Id);
            long componentValue = IDHelper.ElIdValue(componentId);

            if (target == null || !target.IsMultistoryPlacement)
                return isLanding
                    ? $"ПЭ_Л_{stairId}_{componentValue}"
                    : $"ПЭ_{stairId}{componentValue}";

            string prefix = isLanding ? "ПЭ_МЛ" : "ПЭ_М";
            return $"{prefix}_{target.OwnerElementId}_{FormatOptionalElementId(target.PlacementLevelId)}_{componentValue}";
        }

        // =========================
        // НИЗ / ЗАЗОРЫ
        // =========================
        private static double GetMidGapFt(RunTopSearchContext topSearch, Plane undersidePlane, XYZ bottomCenter, XYZ xP, double lenPlan, double widthFt, XYZ yP, bool includeFinish)
        {
            double[] u = new double[] { 0.08, 0.14, 0.20, 0.26, 0.32, 0.38, 0.44, 0.50, 0.56, 0.62, 0.68, 0.74, 0.80, 0.86, 0.92 };

            double sideFactor = 0.35;
            XYZ wSide = yP * (widthFt * sideFactor);

            double maxGap = 0.0;
            int hits = 0;

            for (int i = 0; i < u.Length; i++)
            {
                double t = lenPlan * u[i];
                XYZ c = bottomCenter + xP * t;

                XYZ[] samples =
                {
                    new XYZ(c.X, c.Y, 0) - wSide,
                    new XYZ(c.X, c.Y, 0),
                    new XYZ(c.X, c.Y, 0) + wSide
                };

                for (int s = 0; s < samples.Length; s++)
                {
                    XYZ p = samples[s];

                    if (!TryGetTopZByVerticalIntersect(topSearch, p, includeFinish, out double zTop))
                        continue;

                    double zUnder = GetPlaneZAtXY(undersidePlane, p.X, p.Y);
                    double gap = zTop - zUnder;

                    if (gap > maxGap) maxGap = gap;
                    hits++;
                }
            }

            if (hits == 0) return 0.0;
            if (maxGap < 0.0) maxGap = 0.0;

            return maxGap;
        }

        private static RunTopSearchContext BuildRunTopSearchContext(Document doc, Stairs stairs, StairsRun run)
        {
            var context = new RunTopSearchContext
            {
                MinZ = double.NegativeInfinity,
                MaxZ = double.PositiveInfinity
            };

            AddElementSolids(run, context.RunSolids);

            BoundingBoxXYZ runBox = run.get_BoundingBox(null);
            if (doc == null || runBox == null)
                return context;

            // Используем отделку только как небольшую добавку над найденным верхом текущего марша.
            // Это не должно поднимать тело на всю высоту лестницы.
            AddElementSolids(stairs, context.FinishSolids);

            double xyPaddingFt = MmToInternal(50.0);
            double belowRunFt = MmToInternal(50.0);
            double aboveRunFt = GetMaxFinishThicknessFt();

            context.MinZ = runBox.Min.Z - belowRunFt;
            context.MaxZ = runBox.Max.Z + aboveRunFt;

            Outline outline = new Outline(
                new XYZ(runBox.Min.X - xyPaddingFt, runBox.Min.Y - xyPaddingFt, context.MinZ),
                new XYZ(runBox.Max.X + xyPaddingFt, runBox.Max.Y + xyPaddingFt, context.MaxZ));

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_Parts
            };

            var excludedIds = new HashSet<long> { IDHelper.ElIdValue(run.Id) };
            if (stairs != null)
                excludedIds.Add(IDHelper.ElIdValue(stairs.Id));

            IEnumerable<Element> candidates;
            try
            {
                candidates = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementMulticategoryFilter(categories))
                    .WherePasses(new BoundingBoxIntersectsFilter(outline))
                    .ToElements();
            }
            catch
            {
                return context;
            }

            foreach (Element elem in candidates)
            {
                if (elem == null) continue;
                if (excludedIds.Contains(IDHelper.ElIdValue(elem.Id))) continue;
                if (IsOwnRouteShape(elem)) continue;

                BoundingBoxXYZ bb;
                try
                {
                    bb = elem.get_BoundingBox(null);
                }
                catch
                {
                    continue;
                }

                if (bb == null) continue;
                if (bb.Max.Z < context.MinZ || bb.Min.Z > context.MaxZ) continue;

                AddElementSolids(elem, context.FinishSolids);
            }

            return context;
        }

        private static bool TryGetTopZByVerticalIntersect(RunTopSearchContext topSearch, XYZ pointXY, bool includeFinish, out double zTop)
        {
            zTop = double.NegativeInfinity;

            if (topSearch == null || topSearch.RunSolids == null || topSearch.RunSolids.Count == 0)
                return false;

            double rayPaddingFt = MmToInternal(1000.0);

            double topZ = double.IsPositiveInfinity(topSearch.MaxZ) ? MmToInternal(20000.0) : topSearch.MaxZ + rayPaddingFt;
            double botZ = double.IsNegativeInfinity(topSearch.MinZ) ? -MmToInternal(20000.0) : topSearch.MinZ - rayPaddingFt;
            if (topZ <= botZ) return false;

            XYZ pTop = new XYZ(pointXY.X, pointXY.Y, topZ);
            XYZ pBot = new XYZ(pointXY.X, pointXY.Y, botZ);
            Line line = Line.CreateBound(pTop, pBot);

            if (!TryGetTopZFromSolids(topSearch.RunSolids, line, topSearch.MinZ, topSearch.MaxZ, out double runTopZ))
                return false;

            zTop = runTopZ;

            if (!includeFinish || topSearch.FinishSolids == null || topSearch.FinishSolids.Count == 0)
                return true;

            double finishBottomToleranceFt = MmToInternal(5.0);
            double minFinishZ = runTopZ - finishBottomToleranceFt;
            double maxFinishZ = runTopZ + GetMaxFinishThicknessFt();

            if (TryGetTopZFromSolids(topSearch.FinishSolids, line, minFinishZ, maxFinishZ, out double finishTopZ) && finishTopZ > zTop)
                zTop = finishTopZ;

            return true;
        }

        private static bool TryGetTopZFromSolids(List<Solid> solids, Line line, double minZ, double maxZ, out double topZ)
        {
            topZ = double.NegativeInfinity;
            if (solids == null || solids.Count == 0 || line == null)
                return false;

            double zFilterTolFt = MmToInternal(5.0);

            var opts = new SolidCurveIntersectionOptions();
            opts.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside;

            foreach (Solid s in solids)
            {
                if (s == null || s.Volume < 1e-9) continue;

                SolidCurveIntersection sci;
                try
                {
                    sci = s.IntersectWithCurve(line, opts);
                }
                catch
                {
                    continue;
                }

                if (sci == null) continue;

                int segCount = sci.SegmentCount;
                if (segCount <= 0) continue;

                for (int i = 0; i < segCount; i++)
                {
                    Curve seg = sci.GetCurveSegment(i);
                    if (seg == null) continue;

                    XYZ a = seg.GetEndPoint(0);
                    XYZ b = seg.GetEndPoint(1);

                    TryUseTopPoint(a, minZ, maxZ, zFilterTolFt, ref topZ);
                    TryUseTopPoint(b, minZ, maxZ, zFilterTolFt, ref topZ);
                }
            }

            return !double.IsNegativeInfinity(topZ);
        }

        private static void TryUseTopPoint(XYZ p, double minZ, double maxZ, double zFilterTolFt, ref double topZ)
        {
            if (p == null) return;
            if (!double.IsNegativeInfinity(minZ) && p.Z < minZ - zFilterTolFt) return;
            if (!double.IsPositiveInfinity(maxZ) && p.Z > maxZ + zFilterTolFt) return;
            if (p.Z > topZ) topZ = p.Z;
        }

        private static void AddElementSolids(Element elem, List<Solid> solids)
        {
            if (elem == null || solids == null) return;

            GeometryElement ge;
            try
            {
                ge = elem.get_Geometry(CreateFineGeometryOptions());
            }
            catch
            {
                return;
            }

            if (ge == null) return;

            CollectSolidsRecursive(ge, Transform.Identity, solids);
        }

        private static Options CreateFineGeometryOptions()
        {
            return new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = false,
                IncludeNonVisibleObjects = false
            };
        }

        private static bool IsOwnRouteShape(Element elem)
        {
            DirectShape ds = elem as DirectShape;
            return ds != null && string.Equals(ds.ApplicationId, "KPLN_Tools", StringComparison.Ordinal);
        }

        private static double MmToInternal(double valueMm)
        {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double GetMaxFinishThicknessFt()
        {
            return MmToInternal(80.0);
        }

        private static void CollectSolidsRecursive(GeometryElement ge, Transform tr, List<Solid> solids)
        {
            foreach (GeometryObject go in ge)
            {
                if (go is Solid s)
                {
                    if (s != null && s.Volume > 1e-9)
                    {
                        Solid ts = (tr != null && !tr.IsIdentity) ? SolidUtils.CreateTransformed(s, tr) : s;
                        solids.Add(ts);
                    }
                    continue;
                }

                if (go is GeometryInstance gi)
                {
                    Transform t2 = tr.Multiply(gi.Transform);
                    GeometryElement instGe = gi.GetInstanceGeometry();
                    if (instGe != null)
                        CollectSolidsRecursive(instGe, t2, solids);
                }
            }
        }

        private static bool TryGetBestUndersidePlane(Element elem, out Plane plane)
        {
            plane = null;

            var solids = new List<Solid>();
            AddElementSolids(elem, solids);
            if (solids.Count == 0) return false;

            PlanarFace best = null;
            double bestScore = double.NegativeInfinity;

            foreach (Solid s in solids)
            {
                if (s == null || s.Volume < 1e-9) continue;

                foreach (Face f in s.Faces)
                {
                    PlanarFace pf = f as PlanarFace;
                    if (pf == null) continue;

                    XYZ n = pf.FaceNormal;
                    if (n.Z >= -0.05) continue;

                    double score = (-n.Z) * 10.0 + pf.Area;

                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = pf;
                    }
                }
            }

            if (best == null) return false;

            plane = Plane.CreateByNormalAndOrigin(best.FaceNormal, best.Origin);
            return true;
        }

        private static double GetPlaneZAtXY(Plane plane, double x, double y)
        {
            XYZ n = plane.Normal;
            XYZ p0 = plane.Origin;

            if (Math.Abs(n.Z) < 1e-9)
                return p0.Z;

            double dx = x - p0.X;
            double dy = y - p0.Y;

            return p0.Z - (n.X * dx + n.Y * dy) / n.Z;
        }

        // =========================
        // ГЕОМЕТРИЯ SOLID
        // =========================
        private static Solid BuildPrismFrom8Points(XYZ SL, XYZ SR, XYZ ER, XYZ EL, XYZ SLt, XYZ SRt, XYZ ERt, XYZ ELt)
        {
            TessellatedShapeBuilder tsb = new TessellatedShapeBuilder();
            tsb.OpenConnectedFaceSet(false);

            void AddQuad(XYZ a, XYZ b, XYZ c, XYZ d)
            {
                tsb.AddFace(new TessellatedFace(new List<XYZ> { a, b, c }, ElementId.InvalidElementId));
                tsb.AddFace(new TessellatedFace(new List<XYZ> { a, c, d }, ElementId.InvalidElementId));
            }

            AddQuad(SL, SR, ER, EL);
            AddQuad(SLt, ELt, ERt, SRt);
            AddQuad(SL, EL, ELt, SLt);
            AddQuad(SR, SRt, ERt, ER);
            AddQuad(SL, SLt, SRt, SR);
            AddQuad(EL, ER, ERt, ELt);

            tsb.CloseConnectedFaceSet();

            tsb.Target = TessellatedShapeBuilderTarget.Solid;
            tsb.Fallback = TessellatedShapeBuilderFallback.Abort;

            tsb.Build();

            TessellatedShapeBuilderResult result = tsb.GetBuildResult();
            if (result == null) return null;

            IList<GeometryObject> geom = result.GetGeometricalObjects();
            if (geom == null || geom.Count == 0) return null;

            return geom.OfType<Solid>().FirstOrDefault();
        }

        // =========================
        // ШИРИНА МАРША
        // =========================
        private static RunClearWidthInfo GetRunClearWidthInfo(Document doc, Stairs stairs, StairsRun run, XYZ bottom, XYZ top, XYZ xP, XYZ yP, double lenPlan)
        {
            double nominalWidthFt = GetRunWidthFt(run, bottom, top);
            var result = new RunClearWidthInfo
            {
                WidthFt = nominalWidthFt,
                CenterOffsetFt = 0.0
            };

            if (doc == null || run == null || nominalWidthFt <= 1e-9 || lenPlan <= 1e-9)
                return result;

            BoundingBoxXYZ runBox = run.get_BoundingBox(null);
            if (runBox == null)
                return result;

            double centerY = new XYZ(bottom.X, bottom.Y, 0).DotProduct(yP);
            double clearMinY = centerY - nominalWidthFt / 2.0;
            double clearMaxY = centerY + nominalWidthFt / 2.0;

            double runX0 = new XYZ(bottom.X, bottom.Y, 0).DotProduct(xP);
            double runX1 = new XYZ(top.X, top.Y, 0).DotProduct(xP);
            if (runX1 < runX0)
            {
                double tmp = runX0;
                runX0 = runX1;
                runX1 = tmp;
            }

            double endMarginFt = Math.Min(lenPlan * 0.05, MmToInternal(300.0));
            if (runX1 - runX0 > endMarginFt * 2.0)
            {
                runX0 += endMarginFt;
                runX1 -= endMarginFt;
            }

            var categories = GetClearWidthObstacleCategories();
            if (categories.Count == 0)
                return result;

            double searchPaddingFt = MmToInternal(600.0);
            double searchAboveFt = MmToInternal(1600.0);
            double searchBelowFt = MmToInternal(100.0);

            Outline outline = new Outline(
                new XYZ(runBox.Min.X - searchPaddingFt, runBox.Min.Y - searchPaddingFt, runBox.Min.Z - searchBelowFt),
                new XYZ(runBox.Max.X + searchPaddingFt, runBox.Max.Y + searchPaddingFt, runBox.Max.Z + searchAboveFt));

            var excludedIds = new HashSet<long> { IDHelper.ElIdValue(run.Id) };
            if (stairs != null)
                excludedIds.Add(IDHelper.ElIdValue(stairs.Id));

            IEnumerable<Element> candidates;
            try
            {
                candidates = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementMulticategoryFilter(categories))
                    .WherePasses(new BoundingBoxIntersectsFilter(outline))
                    .ToElements();
            }
            catch
            {
                return result;
            }

            double sideToleranceFt = MmToInternal(150.0);
            double minUsableWidthFt = MmToInternal(300.0);

            foreach (Element elem in candidates)
            {
                if (elem == null) continue;
                if (excludedIds.Contains(IDHelper.ElIdValue(elem.Id))) continue;
                if (IsOwnRouteShape(elem)) continue;

                var solids = new List<Solid>();
                AddElementSolids(elem, solids);
                if (solids.Count == 0) continue;

                foreach (Solid solid in solids)
                {
                    if (!TryGetSolidProjectionRange(solid, xP, yP, out double minX, out double maxX, out double minY, out double maxY))
                        continue;

                    if (maxX < runX0 || minX > runX1)
                        continue;

                    bool fromLeft = maxY <= centerY && maxY > clearMinY - sideToleranceFt && minY < centerY;
                    bool fromRight = minY >= centerY && minY < clearMaxY + sideToleranceFt && maxY > centerY;

                    if (fromLeft)
                        clearMinY = Math.Max(clearMinY, Math.Min(maxY, clearMaxY));

                    if (fromRight)
                        clearMaxY = Math.Min(clearMaxY, Math.Max(minY, clearMinY));
                }
            }

            double clearWidthFt = clearMaxY - clearMinY;
            if (clearWidthFt < minUsableWidthFt || clearWidthFt > nominalWidthFt)
                return result;

            if (nominalWidthFt - clearWidthFt < MmToInternal(1.0))
                return result;

            result.WidthFt = clearWidthFt;
            result.CenterOffsetFt = (clearMinY + clearMaxY) * 0.5 - centerY;
            return result;
        }

        private static List<BuiltInCategory> GetClearWidthObstacleCategories()
        {
            var categories = new List<BuiltInCategory>();

            AddBuiltInCategoryIfDefined(categories, "OST_Railings");
            AddBuiltInCategoryIfDefined(categories, "OST_StairsRailing");
            AddBuiltInCategoryIfDefined(categories, "OST_RailingSystem");
            AddBuiltInCategoryIfDefined(categories, "OST_RailingRail");
            AddBuiltInCategoryIfDefined(categories, "OST_RailingTopRail");
            AddBuiltInCategoryIfDefined(categories, "OST_RailingHandRail");
            AddBuiltInCategoryIfDefined(categories, "OST_RailingSupport");
            AddBuiltInCategoryIfDefined(categories, "OST_Walls");
            AddBuiltInCategoryIfDefined(categories, "OST_Floors");
            AddBuiltInCategoryIfDefined(categories, "OST_Parts");
            AddBuiltInCategoryIfDefined(categories, "OST_GenericModel");

            return categories;
        }

        private static void AddBuiltInCategoryIfDefined(List<BuiltInCategory> categories, string name)
        {
            try
            {
                var bic = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), name);
                if (!categories.Contains(bic))
                    categories.Add(bic);
            }
            catch
            {
            }
        }

        private static bool TryGetSolidProjectionRange(Solid solid, XYZ xP, XYZ yP, out double minX, out double maxX, out double minY, out double maxY)
        {
            minX = double.PositiveInfinity;
            maxX = double.NegativeInfinity;
            minY = double.PositiveInfinity;
            maxY = double.NegativeInfinity;

            if (solid == null || solid.Volume < 1e-9)
                return false;

            bool hasPoints = false;

            foreach (Face face in solid.Faces)
            {
                Mesh mesh;
                try
                {
                    mesh = face.Triangulate();
                }
                catch
                {
                    continue;
                }

                if (mesh == null) continue;

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    if (triangle == null) continue;

                    for (int j = 0; j < 3; j++)
                    {
                        XYZ p = triangle.get_Vertex(j);
                        if (p == null) continue;

                        XYZ pxy = new XYZ(p.X, p.Y, 0);
                        double tx = pxy.DotProduct(xP);
                        double ty = pxy.DotProduct(yP);

                        if (tx < minX) minX = tx;
                        if (tx > maxX) maxX = tx;
                        if (ty < minY) minY = ty;
                        if (ty > maxY) maxY = ty;

                        hasPoints = true;
                    }
                }
            }

            return hasPoints;
        }

        private static double GetRunWidthFt(StairsRun run, XYZ bottom, XYZ top)
        {
            try
            {
                var prop = run.GetType().GetProperty("ActualRunWidth");
                if (prop != null)
                {
                    object v = prop.GetValue(run, null);
                    if (v is double d && d > 1e-9) return d;
                }
            }
            catch { }

            XYZ dirPlan = new XYZ(top.X - bottom.X, top.Y - bottom.Y, 0.0);
            if (dirPlan.GetLength() < 1e-9)
            {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
                return UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);
#else
                return UnitUtils.ConvertToInternalUnits(1000.0, DisplayUnitType.DUT_MILLIMETERS);
#endif
            }

            XYZ xPlan = dirPlan.Normalize();
            XYZ yDir = XYZ.BasisZ.CrossProduct(xPlan);
            if (yDir.GetLength() < 1e-9) yDir = XYZ.BasisY;
            yDir = yDir.Normalize();

            BoundingBoxXYZ bb = run.get_BoundingBox(null);
            if (bb == null)
            {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
                return UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);
#else
                return UnitUtils.ConvertToInternalUnits(1000.0, DisplayUnitType.DUT_MILLIMETERS);
#endif
            }

            var pts = new[]
            {
                new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
            };

            double min = double.PositiveInfinity;
            double max = double.NegativeInfinity;

            foreach (var p in pts)
            {
                double t = p.DotProduct(yDir);
                if (t < min) min = t;
                if (t > max) max = t;
            }

            double w = max - min;
            if (w <= 1e-6)
            {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
                w = UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);
#else
                w = UnitUtils.ConvertToInternalUnits(1000.0, DisplayUnitType.DUT_MILLIMETERS);
#endif
            }

            return w;
        }
    }
}