using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.ExternalCommands.UI;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Interop;
using System.Reflection;
using System.Text.RegularExpressions;
using KPLN_Tools.Common;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_EvacuationRoutes : IExternalCommand
    {
        private const bool UseExperimentalLandingNarrowSection = true;

        private static UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;

        private enum EvacuationRoutesRequestKind
        {
            None,
            SelectElement,
            Build,
            PickAndBuild,
            PickDebugStair,
            PickResizeRoute
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

            public void RequestPickDebugStair(EvacuationRoutesDialog dialog, EvacuationRoutesDialogResult data)
            {
                _dialog = dialog;
                _data = data;
                _elementId = 0;
                _requestKind = EvacuationRoutesRequestKind.PickDebugStair;
            }

            public void RequestPickResizeRoute(EvacuationRoutesDialog dialog)
            {
                _dialog = dialog;
                _data = null;
                _elementId = 0;
                _requestKind = EvacuationRoutesRequestKind.PickResizeRoute;
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

                    if (requestKind == EvacuationRoutesRequestKind.PickDebugStair)
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
                            Finish(dialog, "Debug-отчёт отменён.");
                            return;
                        }

                        EvacuationRoutesDialogResult debugData = new EvacuationRoutesDialogResult(
                            data == null ? 2200 : data.HeightMm,
                            data == null ? 1200 : data.WidthMm,
                            data != null && data.UseRunWidth,
                            data == null || data.ConsiderRailings,
                            data == null || data.RoundRunWidthDownTo5Mm,
                            true,
                            data != null && data.AddToEvacuationWorkset,
                            data == null ? (int?)null : data.EvacuationWorksetId,
                            pickedId.Value,
                            null,
                            data == null ? null : data.UseRunWidthByElementId);

                        string path = _owner.SaveStairDebugReportToDesktop(pickedId.Value, debugData);
                        Finish(dialog, $"Debug-отчёт сохранён: {path}");
                        return;
                    }

                    if (requestKind == EvacuationRoutesRequestKind.PickResizeRoute)
                    {
                        long? pickedRouteId;
                        HideForPick(dialog);
                        try
                        {
                            pickedRouteId = PickEvacuationRouteElementId(app, _owner.doc);
                        }
                        finally
                        {
                            RestoreAfterPick(dialog);
                        }

                        if (!pickedRouteId.HasValue)
                        {
                            Finish(dialog, "Изменение габаритов отменено.");
                            return;
                        }

                        EvacuationRoutesCheckRequest pickedRoute = _owner.CreateRouteCheckRequestForRoute(pickedRouteId.Value);
                        RouteDimensionInfo dimensions = GetRouteDimensionInfo(_owner.doc, pickedRoute.RouteElementId, pickedRoute.ComponentElementId);
                        if (dimensions.LengthMm <= 0 || dimensions.WidthMm <= 0 || dimensions.HeightMm <= 0)
                        {
                            Finish(dialog, $"Не удалось определить габариты пути ID {pickedRouteId.Value}.");
                            return;
                        }

                        EvacuationRoutesResizeRequest pickedResizeRequest;
                        if (!TryGetResizeRequestFromDialog(dialog, pickedRoute, dimensions, out pickedResizeRequest))
                        {
                            Finish(dialog, "Изменение габаритов отменено.");
                            return;
                        }

                        RouteEditResult resizeResult = _owner.ResizeRouteShape(pickedResizeRequest);
                        Finish(dialog, resizeResult.Message);
                        ShowRouteCheck(dialog, resizeResult.CheckReport);
                        UpdateDimensions(dialog, resizeResult);
                        if (resizeResult.IsFixed)
                            MarkFixed(dialog, resizeResult.StairElementId);
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
                            data.ConsiderRailings,
                            data.RoundRunWidthDownTo5Mm,
                            true,
                            data.AddToEvacuationWorkset,
                            data.EvacuationWorksetId,
                            pickedId.Value,
                            null,
                            data.UseRunWidthByElementId);

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

            private static void ShowRouteCheck(EvacuationRoutesDialog dialog, string text)
            {
                if (dialog == null || string.IsNullOrWhiteSpace(text))
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.ShowRouteCheckResult(text)));
            }

            private static void MarkFixed(EvacuationRoutesDialog dialog, long stairId)
            {
                if (dialog == null || stairId <= 0)
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.MarkStairFixed(stairId)));
            }

            private static void UpdateDimensions(EvacuationRoutesDialog dialog, RouteEditResult result)
            {
                if (dialog == null || result == null || !result.HasDimensions || result.RouteElementId <= 0)
                    return;

                dialog.Dispatcher.BeginInvoke(new Action(() => dialog.UpdateRouteDimensions(
                    result.RouteElementId,
                    result.LengthMm,
                    result.WidthMm,
                    result.HeightMm)));
            }

            private static bool TryGetResizeRequestFromDialog(EvacuationRoutesDialog dialog, EvacuationRoutesCheckRequest route, RouteDimensionInfo dimensions, out EvacuationRoutesResizeRequest request)
            {
                request = null;
                if (dialog == null || route == null)
                    return false;

                EvacuationRoutesResizeRequest localRequest = null;
                bool accepted = false;
                dialog.Dispatcher.Invoke(new Action(() =>
                {
                    accepted = dialog.TryCreateResizeRequestForRoute(
                        route.RouteElementId,
                        route.StairElementId,
                        route.ComponentElementId,
                        dimensions.LengthMm,
                        dimensions.WidthMm,
                        dimensions.HeightMm,
                        out localRequest);
                }));

                request = localRequest;
                return accepted && request != null;
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

        private sealed class EvacuationRouteSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => IsEvacuationRouteShape(elem);
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        private sealed class RunRouteBodyInfo
        {
            public ElementId RunId;
            public int StairsId;

            public double WidthFt;
            public double HeightFt;

            public XYZ XDirPlan; 
            public XYZ YDirPlan; 
            public EndFace BottomEnd; 
            public EndFace TopEnd; 

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
            public double ClearMinY;
            public double ClearMaxY;
            public bool HasRailingBoundary;
            public bool HasLeftRailingBoundary;
            public bool HasRightRailingBoundary;
        }

        private struct ProjectionRange2D
        {
            public double MinX;
            public double MaxX;
            public double MinY;
            public double MaxY;
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

        private sealed class ExistingRouteCheckResult
        {
            public int ExpectedCount;
            public int FoundCount;

            public bool HasAny => FoundCount > 0;
            public bool IsComplete => ExpectedCount > 0 && FoundCount >= ExpectedCount;
            public bool IsPartial => FoundCount > 0 && !IsComplete;

            public void Add(ExistingRouteCheckResult other)
            {
                if (other == null)
                    return;

                ExpectedCount += other.ExpectedCount;
                FoundCount += other.FoundCount;
            }
        }

        private sealed class RouteCheckResult
        {
            public bool HasIntersections;
            public string ReportText;
        }

        private sealed class RouteEditResult
        {
            public long StairElementId;
            public long RouteElementId;
            public bool HasDimensions;
            public double LengthMm;
            public double WidthMm;
            public double HeightMm;
            public bool IsFixed;
            public string Message;
            public string CheckReport;
        }

        private struct RouteDimensionInfo
        {
            public double LengthMm;
            public double WidthMm;
            public double HeightMm;
            public XYZ XDir;
            public XYZ YDir;
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
            public XYZ BL;
            public XYZ BR;  
            public XYZ TR; 
            public XYZ TL; 

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

        private struct LocalRect2D
        {
            public double MinX;
            public double MaxX;
            public double MinY;
            public double MaxY;
            public string Name;
            public string Group;

            public LocalRect2D(double minX, double maxX, double minY, double maxY, string name, string group = null)
            {
                MinX = Math.Min(minX, maxX);
                MaxX = Math.Max(minX, maxX);
                MinY = Math.Min(minY, maxY);
                MaxY = Math.Max(minY, maxY);
                Name = name ?? "";
                Group = string.IsNullOrWhiteSpace(group) ? Name : group;
            }
        }

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
                },
                () =>
                {
                    handler.RequestPickResizeRoute(dlg);
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

            HashSet<string> existingRouteAppDataIds = CollectExistingRouteAppDataIds(doc);

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

                var item = new EvacuationRoutesStairListItem
                {
                    ElementId = IDHelper.ElIdValue(multistory.Id),
                    Kind = "Многоэтажная",
                    Name = GetElementDisplayName(multistory),
                    TypeName = GetElementTypeDisplayName(doc, multistory),
                    WorksetName = GetElementWorksetName(doc, multistory),
                    RunCount = GetStairRunCount(standardStairs),
                    LandingCount = GetStairLandingCount(standardStairs),
                    NestedCount = placementCount > 0 ? placementCount : nestedIds.Count,
                    ConnectedLevelCount = connectedLevelCount,
                    NestedStairIds = nestedIds
                };

                ApplyExistingRouteState(item, GetExistingRouteStateForMultistory(doc, multistory, existingRouteAppDataIds));
                result.Add(item);

                foreach (long nestedId in nestedIds)
                    seenStairs.Add(nestedId);
            }

            foreach (Stairs stairs in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .OfType<Stairs>())
            {
                AddStairListItemIfNew(doc, stairs, result, seenStairs, null, existingRouteAppDataIds);
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

        private static void AddStairListItemIfNew(Document doc, Stairs stairs, List<EvacuationRoutesStairListItem> result, HashSet<long> seenStairs, long? parentMultistoryId, HashSet<string> existingRouteAppDataIds)
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
                WorksetName = GetElementWorksetName(doc, stairs),
                RunCount = GetStairRunCount(stairs),
                LandingCount = GetStairLandingCount(stairs),
                NestedCount = 0,
                ParentMultistoryId = parentMultistoryId
            };

            ApplyExistingRouteState(item, GetExistingRouteStateForTarget(CreateOrdinaryRouteBuildTarget(stairs), existingRouteAppDataIds));
            result.Add(item);
        }

        private static void ApplyExistingRouteState(EvacuationRoutesStairListItem item, ExistingRouteCheckResult routeState)
        {
            if (item == null || routeState == null || !routeState.HasAny)
                return;

            item.HasExistingRoute = true;
            item.IsIncluded = false;

            if (routeState.IsComplete)
            {
                item.ExistingRouteStatus = EvacuationRoutesStatus.Built;
                item.SetStatus(EvacuationRoutesStatus.Built, "Построено", null);
            }
            else
            {
                item.ExistingRouteStatus = EvacuationRoutesStatus.PartialBuilt;
                item.SetStatus(EvacuationRoutesStatus.PartialBuilt, "Частично", null);
            }
        }

        private static HashSet<string> CollectExistingRouteAppDataIds(Document doc)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            if (doc == null)
                return result;

            try
            {
                foreach (DirectShape ds in new FilteredElementCollector(doc)
                    .OfClass(typeof(DirectShape))
                    .WhereElementIsNotElementType()
                    .OfType<DirectShape>())
                {
                    if (!string.Equals(ds.ApplicationId, "KPLN_Tools", StringComparison.Ordinal))
                        continue;

                    if (!string.IsNullOrWhiteSpace(ds.ApplicationDataId))
                        result.Add(ds.ApplicationDataId);
                }
            }
            catch
            {
            }

            return result;
        }

        private static ExistingRouteCheckResult GetExistingRouteStateForMultistory(Document doc, MultistoryStairs multistory, HashSet<string> existingRouteAppDataIds)
        {
            var result = new ExistingRouteCheckResult();
            if (doc == null || multistory == null)
                return result;

            foreach (RouteBuildTarget target in GetMultistoryRouteBuildTargets(doc, multistory))
                result.Add(GetExistingRouteStateForTarget(target, existingRouteAppDataIds));

            return result;
        }

        private static ExistingRouteCheckResult GetExistingRouteStateForTarget(RouteBuildTarget target, HashSet<string> existingRouteAppDataIds)
        {
            var result = new ExistingRouteCheckResult();
            if (target == null || target.Stairs == null)
                return result;

            foreach (ElementId componentId in GetRouteComponentIds(target.Stairs))
            {
                result.ExpectedCount++;

                if (existingRouteAppDataIds == null || existingRouteAppDataIds.Count == 0)
                    continue;

                string appDataId = CreateRouteAppDataId(target, componentId);
                if (!string.IsNullOrWhiteSpace(appDataId) && existingRouteAppDataIds.Contains(appDataId))
                {
                    result.FoundCount++;
                    continue;
                }

                string componentOnly = IDHelper.ElIdValue(componentId).ToString();
                bool canUseLegacyComponentKey = target == null || !target.IsMultistoryPlacement || Math.Abs(target.VerticalOffsetFt) < 1e-9;
                if (canUseLegacyComponentKey && existingRouteAppDataIds.Contains(componentOnly))
                    result.FoundCount++;
            }

            return result;
        }

        private static IEnumerable<ElementId> GetRouteComponentIds(Stairs stairs)
        {
            if (stairs == null)
                yield break;

            ICollection<ElementId> runIds = null;
            ICollection<ElementId> landingIds = null;

            try { runIds = stairs.GetStairsRuns(); } catch { }
            try { landingIds = stairs.GetStairsLandings(); } catch { }

            foreach (ElementId id in runIds ?? new List<ElementId>())
            {
                if (id != null && id != ElementId.InvalidElementId)
                    yield return id;
            }

            foreach (ElementId id in landingIds ?? new List<ElementId>())
            {
                if (id != null && id != ElementId.InvalidElementId)
                    yield return id;
            }
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
                debugLog.Add($"Настройки: HeightMm={data.HeightMm}; WidthMm={data.WidthMm}; UseRunWidth={data.UseRunWidth}; ConsiderRailings={data.ConsiderRailings}; RoundRunWidthDownTo5Mm={data.RoundRunWidthDownTo5Mm}; AddToEvacuationWorkset={data.AddToEvacuationWorkset}; WorksetId={(data.EvacuationWorksetId.HasValue ? data.EvacuationWorksetId.Value.ToString() : "null")}; SelectedElementId={(data.SelectedElementId.HasValue ? data.SelectedElementId.Value.ToString() : "null")}");
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

                try { doc.Regenerate(); } catch { }
                t.Commit();
            }

            TryRefreshActiveView();

            AddMultistoryAggregateUpdates(result, buildResults);
            FillOperationReportLines(result, data, targets, intersectionReports, debugLog, buildResults);

            return result;
        }

        private void TryRefreshActiveView()
        {
            try
            {
                uidoc?.RefreshActiveView();
            }
            catch
            {
            }
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

        private static bool ShouldUseRunWidth(EvacuationRoutesDialogResult data, RouteBuildTarget target)
        {
            if (data == null)
                return false;

            Dictionary<long, bool> byElementId = data.UseRunWidthByElementId;
            if (byElementId != null && target != null)
            {
                bool value;
                if (target.OwnerElementId > 0 && byElementId.TryGetValue(target.OwnerElementId, out value))
                    return value;

                if (target.StandardStairsId > 0 && byElementId.TryGetValue(target.StandardStairsId, out value))
                    return value;
            }

            return data.UseRunWidth;
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
            result.ProblemGroups = BuildProblemGroups(doc, buildResults, intersectionReports);
        }

        private static List<EvacuationRoutesProblemGroup> BuildProblemGroups(Document doc, List<RouteBuildTargetResult> buildResults, List<RouteIntersectionReportItem> intersectionReports)
        {
            var grouped = new SortedDictionary<long, EvacuationRoutesProblemGroup>();

            Func<long, EvacuationRoutesProblemGroup> getGroup = stairId =>
            {
                if (!grouped.TryGetValue(stairId, out EvacuationRoutesProblemGroup group))
                {
                    group = new EvacuationRoutesProblemGroup { StairElementId = stairId };
                    grouped[stairId] = group;
                }

                return group;
            };

            foreach (RouteBuildTargetResult buildResult in buildResults ?? new List<RouteBuildTargetResult>())
            {
                RouteBuildTarget target = buildResult?.Target;
                if (target == null)
                    continue;

                long stairId = target.OwnerElementId > 0 ? target.OwnerElementId : target.StandardStairsId;
                EvacuationRoutesProblemGroup group = getGroup(stairId);

                foreach (int runId in buildResult.FailedRuns ?? new List<int>())
                {
                    group.Items.Add(new EvacuationRoutesProblemItem
                    {
                        ComponentKind = "Марш",
                        ComponentElementId = runId,
                        Message = "не построился"
                    });
                }

                foreach (int landingId in buildResult.FailedLandings ?? new List<int>())
                {
                    group.Items.Add(new EvacuationRoutesProblemItem
                    {
                        ComponentKind = "Площадка",
                        ComponentElementId = landingId,
                        Message = "не построилась"
                    });
                }

                if (!buildResult.Ok && (buildResult.FailedRuns == null || buildResult.FailedRuns.Count == 0) && (buildResult.FailedLandings == null || buildResult.FailedLandings.Count == 0))
                {
                    group.Items.Add(new EvacuationRoutesProblemItem
                    {
                        ComponentKind = target.IsMultistoryPlacement ? "Размещение" : "Лестница",
                        ComponentElementId = target.StandardStairsId,
                        Message = "не построилось"
                    });
                }
            }

            foreach (RouteIntersectionReportItem report in intersectionReports ?? new List<RouteIntersectionReportItem>())
            {
                if (report == null || report.Targets == null || report.Targets.Count == 0)
                    continue;

                long stairId = report.OwnerElementId > 0 ? report.OwnerElementId : 0;
                EvacuationRoutesProblemGroup group = getGroup(stairId);

                RouteDimensionInfo dimensions = GetRouteDimensionInfo(doc, report.RouteElementId, report.ComponentElementId);

                group.Items.Add(new EvacuationRoutesProblemItem
                {
                    ComponentKind = string.IsNullOrWhiteSpace(report.ComponentKind) ? "Элемент" : report.ComponentKind,
                    ComponentElementId = report.ComponentElementId,
                    RouteElementId = report.RouteElementId,
                    Message = "пересечение",
                    CurrentLengthMm = dimensions.LengthMm,
                    CurrentWidthMm = dimensions.WidthMm,
                    CurrentHeightMm = dimensions.HeightMm,
                    Targets = report.Targets
                        .OrderBy(x => x.ElementId)
                        .Select(x => new EvacuationRoutesProblemTarget
                        {
                            ElementId = x.ElementId,
                            LinkInstanceId = x.LinkInstanceId,
                            DisplayText = FormatIntersectionTargetShort(x)
                        })
                        .ToList()
                });
            }

            return grouped.Values
                .Where(x => x != null && x.Items != null && x.Items.Count > 0)
                .ToList();
        }

        private static RouteDimensionInfo GetRouteDimensionInfo(Document doc, long routeElementId, long componentElementId)
        {
            var result = new RouteDimensionInfo
            {
                XDir = XYZ.BasisX,
                YDir = XYZ.BasisY
            };

            if (doc == null || routeElementId <= 0)
                return result;

            Element routeElement = doc.GetElement(IDHelper.CreateElementId(routeElementId));
            if (routeElement == null)
                return result;

            componentElementId = ResolveRouteComponentElementId(doc, routeElement, componentElementId);

            XYZ xDir;
            XYZ yDir;
            GetRouteAxes(doc, componentElementId, routeElement, out xDir, out yDir);
            result.XDir = xDir;
            result.YDir = yDir;

            var solids = new List<Solid>();
            AddElementSolids(routeElement, solids);
            List<XYZ> vertices = CollectSolidVertices(GetValidSolids(solids));
            if (vertices.Count == 0)
                return result;

            double minX = double.PositiveInfinity;
            double maxX = double.NegativeInfinity;
            double minY = double.PositiveInfinity;
            double maxY = double.NegativeInfinity;

            foreach (XYZ p in vertices)
            {
                XYZ pxy = new XYZ(p.X, p.Y, 0.0);
                double x = pxy.DotProduct(xDir);
                double y = pxy.DotProduct(yDir);
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
            }

            if (!double.IsInfinity(minX) && !double.IsInfinity(maxX))
                result.LengthMm = IDHelper.ConvertInternalToMm(maxX - minX);

            if (!double.IsInfinity(minY) && !double.IsInfinity(maxY))
                result.WidthMm = IDHelper.ConvertInternalToMm(maxY - minY);

            double heightFt = EstimateVerticalThicknessFt(vertices);
            if (heightFt <= 1e-9)
                heightFt = vertices.Max(x => x.Z) - vertices.Min(x => x.Z);

            if (heightFt > 1e-9)
                result.HeightMm = IDHelper.ConvertInternalToMm(heightFt);

            return result;
        }

        private static long ResolveRouteComponentElementId(Document doc, Element routeElement, long componentElementId)
        {
            if (doc != null && componentElementId > 0)
            {
                try
                {
                    if (doc.GetElement(IDHelper.CreateElementId(componentElementId)) != null)
                        return componentElementId;
                }
                catch
                {
                }
            }

            DirectShape directShape = routeElement as DirectShape;
            string appDataId = directShape == null ? null : directShape.ApplicationDataId;
            if (string.IsNullOrWhiteSpace(appDataId))
                return componentElementId;

            MatchCollection matches = Regex.Matches(appDataId, @"\d+");
            if (matches.Count == 0)
                return componentElementId;

            string raw = matches[matches.Count - 1].Value;
            if (!long.TryParse(raw, out long parsedId) || parsedId <= 0)
                return componentElementId;

            if (doc != null)
            {
                try
                {
                    if (doc.GetElement(IDHelper.CreateElementId(parsedId)) == null)
                        return componentElementId;
                }
                catch
                {
                    return componentElementId;
                }
            }

            return parsedId;
        }

        private static void GetRouteAxes(Document doc, long componentElementId, Element routeElement, out XYZ xDir, out XYZ yDir)
        {
            xDir = XYZ.BasisX;
            yDir = XYZ.BasisY;

            try
            {
                Element component = doc?.GetElement(IDHelper.CreateElementId(componentElementId));
                StairsRun run = component as StairsRun;
                if (run != null)
                {
                    CurveLoop path = run.GetStairsPath();
                    List<Curve> curves = path == null ? new List<Curve>() : path.ToList();
                    if (curves.Count > 0)
                    {
                        XYZ a = curves.First().GetEndPoint(0);
                        XYZ b = curves.Last().GetEndPoint(1);
                        XYZ d = new XYZ(b.X - a.X, b.Y - a.Y, 0.0);
                        if (d.GetLength() > 1e-9)
                        {
                            xDir = d.Normalize();
                            yDir = XYZ.BasisZ.CrossProduct(xDir).Normalize();
                            return;
                        }
                    }
                }

                StairsLanding landing = component as StairsLanding;
                if (landing != null && TryGetLongestHorizontalEdgeDirection(landing, out XYZ landingEdgeDir))
                {
                    xDir = landingEdgeDir;
                    yDir = XYZ.BasisZ.CrossProduct(xDir).Normalize();
                    return;
                }
            }
            catch
            {
            }

            if (TryGetLongestHorizontalEdgeDirection(routeElement, out XYZ edgeDir))
            {
                xDir = edgeDir;
                yDir = XYZ.BasisZ.CrossProduct(xDir).Normalize();
            }
        }

        private static bool TryGetLongestHorizontalEdgeDirection(Element elem, out XYZ xDir)
        {
            xDir = XYZ.BasisX;
            var solids = new List<Solid>();
            AddElementSolids(elem, solids);

            double best = 0.0;
            XYZ bestDir = null;
            double zTol = MmToInternal(2.0);

            foreach (Solid solid in GetValidSolids(solids))
            {
                foreach (Face face in solid.Faces)
                {
                    IList<CurveLoop> loops = null;
                    try { loops = face.GetEdgesAsCurveLoops(); } catch { }
                    foreach (CurveLoop loop in loops ?? new List<CurveLoop>())
                    {
                        foreach (Curve curve in loop)
                        {
                            XYZ a = curve.GetEndPoint(0);
                            XYZ b = curve.GetEndPoint(1);
                            if (Math.Abs(a.Z - b.Z) > zTol)
                                continue;

                            XYZ d = new XYZ(b.X - a.X, b.Y - a.Y, 0.0);
                            double len = d.GetLength();
                            if (len > best)
                            {
                                best = len;
                                bestDir = d.Normalize();
                            }
                        }
                    }
                }
            }

            if (bestDir == null)
                return false;

            xDir = bestDir;
            return true;
        }

        private static List<XYZ> CollectSolidVertices(IEnumerable<Solid> solids)
        {
            var vertices = new List<XYZ>();
            foreach (Solid solid in solids ?? Enumerable.Empty<Solid>())
            {
                if (solid == null || solid.Volume <= 1e-9)
                    continue;

                foreach (Face face in solid.Faces)
                {
                    Mesh mesh = null;
                    try { mesh = face.Triangulate(); } catch { }
                    if (mesh == null)
                        continue;

                    for (int i = 0; i < mesh.NumTriangles; i++)
                    {
                        MeshTriangle triangle = mesh.get_Triangle(i);
                        vertices.Add(triangle.get_Vertex(0));
                        vertices.Add(triangle.get_Vertex(1));
                        vertices.Add(triangle.get_Vertex(2));
                    }
                }
            }

            return vertices;
        }

        private static double EstimateVerticalThicknessFt(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count == 0)
                return 0.0;

            double tol = MmToInternal(1.0);
            var spans = new Dictionary<string, Tuple<double, double>>();

            foreach (XYZ p in vertices)
            {
                string key = $"{Math.Round(p.X / tol)}:{Math.Round(p.Y / tol)}";
                if (!spans.TryGetValue(key, out Tuple<double, double> span))
                {
                    spans[key] = Tuple.Create(p.Z, p.Z);
                    continue;
                }

                spans[key] = Tuple.Create(Math.Min(span.Item1, p.Z), Math.Max(span.Item2, p.Z));
            }

            List<double> values = spans.Values
                .Select(x => x.Item2 - x.Item1)
                .Where(x => x > tol)
                .OrderBy(x => x)
                .ToList();

            if (values.Count > 0)
                return values[values.Count / 2];

            double minZ = vertices.Min(x => x.Z);
            double maxZ = vertices.Max(x => x.Z);
            return Math.Max(0.0, maxZ - minZ);
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
    
        private static string FormatOptionalElementId(ElementId id)
        {
            return id == null || id == ElementId.InvalidElementId
                ? "нет"
                : IDHelper.ElIdValue(id).ToString();
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

        private static long? PickEvacuationRouteElementId(UIApplication uiapp, Document doc)
        {
            try
            {
                var uidoc = uiapp.ActiveUIDocument;
                Reference r = uidoc.Selection.PickObject(ObjectType.Element, new EvacuationRouteSelectionFilter(), "Выберите построенный путь эвакуации ПЭ (Esc — Отмена)");
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

        private string SaveStairDebugReportToDesktop(long selectedElementId, EvacuationRoutesDialogResult data)
        {
            if (doc == null)
                throw new InvalidOperationException("Документ Revit недоступен для debug-отчёта.");

            Element selected = doc.GetElement(IDHelper.CreateElementId(selectedElementId));
            if (selected == null)
                throw new InvalidOperationException($"Элемент ID {selectedElementId} не найден.");

            var lines = new List<string>
            {
                "KPLN. Пути эвакуации — DEBUG-отчёт по выбранной лестнице",
                $"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                $"Документ: {doc.Title}",
                $"Настройки окна: HeightMm={(data == null ? 0 : data.HeightMm)}; WidthMm={(data == null ? 0 : data.WidthMm)}; UseRunWidth={(data != null && data.UseRunWidth)}; ConsiderRailings={(data != null && data.ConsiderRailings)}; RoundRunWidthDownTo5Mm={(data != null && data.RoundRunWidthDownTo5Mm)}",
                ""
            };

            AddDebugElementBlock(lines, "ВЫБРАННЫЙ ЭЛЕМЕНТ", selected, includeParameters: true);

            List<RouteBuildTarget> targets = GetDebugRouteBuildTargets(doc, selected);
            lines.Add("");
            lines.Add($"Целей обработки найдено: {targets.Count}");

            foreach (RouteBuildTarget target in targets)
            {
                if (target == null || target.Stairs == null)
                    continue;

                lines.Add("");
                lines.Add("============================================================");
                lines.Add($"TARGET: {target.DisplayName}");
                lines.Add($"OwnerElementId={target.OwnerElementId}");
                lines.Add($"StandardStairsId={target.StandardStairsId}");
                lines.Add($"PlacementLevelId={FormatOptionalElementId(target.PlacementLevelId)}");
                lines.Add($"VerticalOffset={FormatFtMm(target.VerticalOffsetFt)}");
                lines.Add($"EffectiveUseRunWidth={ShouldUseRunWidth(data, target)}");

                AddDebugStairComponents(lines, target);
                AddDebugExistingRouteShapes(lines, target);
                AddDebugDryRun(lines, target, data);
            }

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"KPLN_EvacuationRoutes_Debug_Stair_{selectedElementId}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string path = System.IO.Path.Combine(desktop, fileName);
            System.IO.File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
            return path;
        }

        private static List<RouteBuildTarget> GetDebugRouteBuildTargets(Document doc, Element selected)
        {
            if (doc == null || selected == null)
                return new List<RouteBuildTarget>();

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

        private static void AddDebugStairComponents(List<string> lines, RouteBuildTarget target)
        {
            Stairs stairs = target == null ? null : target.Stairs;
            if (lines == null || stairs == null)
                return;

            lines.Add("");
            AddDebugElementBlock(lines, "ЛЕСТНИЦА", stairs, includeParameters: true);
            AddDebugReflectedProperties(lines, stairs, "  StairProperty", new[]
            {
                "ActualRiserHeight", "ActualTreadDepth", "ActualRisersNumber", "ActualTreadsNumber",
                "DesiredRisersNumber", "BaseElevation", "TopElevation", "Height", "NumberOfStories",
                "MultistoryStairsId"
            });

            ICollection<ElementId> runIds = null;
            ICollection<ElementId> landingIds = null;
            try { runIds = stairs.GetStairsRuns(); } catch { }
            try { landingIds = stairs.GetStairsLandings(); } catch { }

            lines.Add($"  RunIds={FormatElementIds(runIds)}");
            lines.Add($"  LandingIds={FormatElementIds(landingIds)}");

            foreach (ElementId runId in runIds ?? new List<ElementId>())
            {
                StairsRun run = stairs.Document.GetElement(runId) as StairsRun;
                lines.Add("");
                AddDebugElementBlock(lines, $"МАРШ ID {IDHelper.ElIdValue(runId)}", run, includeParameters: true);
                AddDebugReflectedProperties(lines, run, "  RunProperty", new[]
                {
                    "ActualRunWidth", "ActualRiserHeight", "ActualTreadDepth", "ActualRisersNumber",
                    "ActualTreadsNumber", "BaseElevation", "TopElevation", "Height"
                });
                AddDebugRunPath(lines, run);
                AddDebugExpectedRoute(lines, target, runId, isLanding: false);
            }

            foreach (ElementId landingId in landingIds ?? new List<ElementId>())
            {
                StairsLanding landing = stairs.Document.GetElement(landingId) as StairsLanding;
                lines.Add("");
                AddDebugElementBlock(lines, $"ПЛОЩАДКА ID {IDHelper.ElIdValue(landingId)}", landing, includeParameters: true);
                AddDebugReflectedProperties(lines, landing, "  LandingProperty", new[] { "Height", "Thickness", "BaseElevation", "TopElevation" });
                AddDebugSolidStats(lines, landing, "  LandingSolids");
                AddDebugExpectedRoute(lines, target, landingId, isLanding: true);
            }
        }

        private static void AddDebugExpectedRoute(List<string> lines, RouteBuildTarget target, ElementId componentId, bool isLanding)
        {
            if (lines == null || target == null || componentId == null)
                return;

            string appDataId = CreateRouteAppDataId(target, componentId);
            string routeName = CreateRouteName(target, componentId, isLanding);
            lines.Add($"  ExpectedRouteName={routeName}");
            lines.Add($"  ExpectedApplicationDataId={appDataId}");
        }

        private void AddDebugExistingRouteShapes(List<string> lines, RouteBuildTarget target)
        {
            if (lines == null || doc == null || target == null || target.Stairs == null)
                return;

            lines.Add("");
            lines.Add("ПОСТРОЕННЫЕ ПУТИ ЭВАКУАЦИИ ПО КОМПОНЕНТАМ");

            foreach (Tuple<ElementId, string, bool> component in GetDebugRouteComponents(target.Stairs))
            {
                ElementId componentId = component.Item1;
                string kind = component.Item2;
                bool isLanding = component.Item3;
                string appDataId = CreateRouteAppDataId(target, componentId);
                DirectShape routeShape = FindExistingRouteShape(doc, "KPLN_Tools", appDataId);

                lines.Add("");
                lines.Add($"  {kind} ID {IDHelper.ElIdValue(componentId)} | appDataId={appDataId}");

                if (routeShape == null)
                {
                    lines.Add("    RouteShape=не найден");
                    continue;
                }

                lines.Add($"    RouteShapeId={IDHelper.ElIdValue(routeShape.Id)}");
                lines.Add($"    RouteName={routeShape.Name}");
                lines.Add($"    RouteBoundingBox={FormatBoundingBox(SafeGetBoundingBox(routeShape))}");

                RouteDimensionInfo dimensions = GetRouteDimensionInfo(doc, IDHelper.ElIdValue(routeShape.Id), IDHelper.ElIdValue(componentId));
                lines.Add($"    RouteDimensions: Length={dimensions.LengthMm:0.#} mm; Width={dimensions.WidthMm:0.#} mm; Height={dimensions.HeightMm:0.#} mm");
                AddDebugRouteIntersections(lines, target, componentId, kind, routeShape);
            }
        }

        private void AddDebugDryRun(List<string> lines, RouteBuildTarget target, EvacuationRoutesDialogResult data)
        {
            if (lines == null || doc == null || target == null || target.Stairs == null)
                return;

            lines.Add("");
            lines.Add("DRY-RUN ПОСТРОЕНИЯ (rollback, модель не меняется)");

            var debugLog = new RouteDebugLog { Enabled = true };
            var reports = new List<RouteIntersectionReportItem>();

            Transaction tx = null;
            try
            {
                tx = new Transaction(doc, "KPLN: DEBUG dry-run путей эвакуации");
                tx.Start();

                bool ok = TryCreateRouteBodyOnStair(
                    doc,
                    target,
                    data,
                    reports,
                    debugLog,
                    out int createdRuns,
                    out int createdLandings,
                    out List<int> failedRuns,
                    out List<int> failedLandings);

                lines.Add($"DryRunResult ok={ok}; createdRuns={createdRuns}; createdLandings={createdLandings}; failedRuns={string.Join(", ", failedRuns ?? new List<int>())}; failedLandings={string.Join(", ", failedLandings ?? new List<int>())}");
            }
            catch (Exception ex)
            {
                lines.Add($"DryRun ERROR: {ex}");
            }
            finally
            {
                try
                {
                    if (tx != null && tx.GetStatus() == TransactionStatus.Started)
                        tx.RollBack();
                }
                catch
                {
                }
            }

            lines.Add("");
            lines.Add("DRY-RUN LOG:");
            foreach (string line in debugLog.Lines ?? new List<string>())
                lines.Add("  " + line);

            lines.Add("");
            lines.Add("DRY-RUN INTERSECTIONS:");
            AddDebugIntersectionReports(lines, reports, "  ");
        }

        private void AddDebugRouteIntersections(List<string> lines, RouteBuildTarget target, ElementId componentId, string kind, DirectShape routeShape)
        {
            if (lines == null || routeShape == null)
                return;

            var solids = new List<Solid>();
            AddElementSolids(routeShape, solids);
            solids = GetValidSolids(solids);

            var reports = new List<RouteIntersectionReportItem>();
            AddRouteIntersectionReport(
                doc,
                solids,
                routeShape,
                string.IsNullOrWhiteSpace(routeShape.Name) ? $"Путь ID {IDHelper.ElIdValue(routeShape.Id)}" : routeShape.Name,
                GetStairAndComponentIds(target == null ? null : target.Stairs),
                reports,
                null,
                target,
                componentId,
                kind);

            AddDebugIntersectionReports(lines, reports, "    ");
        }

        private static void AddDebugIntersectionReports(List<string> lines, List<RouteIntersectionReportItem> reports, string indent)
        {
            indent = indent ?? "";
            if (lines == null)
                return;

            if (reports == null || reports.Count == 0)
            {
                lines.Add(indent + "Intersections=нет");
                return;
            }

            foreach (RouteIntersectionReportItem report in reports)
            {
                lines.Add($"{indent}{report.ComponentKind} ID {report.ComponentElementId}; routeId={report.RouteElementId}; targets={(report.Targets == null ? 0 : report.Targets.Count)}");
                foreach (RouteIntersectionTarget target in report.Targets ?? new List<RouteIntersectionTarget>())
                    lines.Add(indent + "- " + FormatIntersectionTargetShort(target));
            }
        }

        private static List<Tuple<ElementId, string, bool>> GetDebugRouteComponents(Stairs stairs)
        {
            var result = new List<Tuple<ElementId, string, bool>>();
            if (stairs == null)
                return result;

            try
            {
                foreach (ElementId id in stairs.GetStairsRuns() ?? new List<ElementId>())
                    result.Add(Tuple.Create(id, "Марш", false));
            }
            catch
            {
            }

            try
            {
                foreach (ElementId id in stairs.GetStairsLandings() ?? new List<ElementId>())
                    result.Add(Tuple.Create(id, "Площадка", true));
            }
            catch
            {
            }

            return result;
        }

        private static void AddDebugRunPath(List<string> lines, StairsRun run)
        {
            if (lines == null)
                return;

            if (run == null)
            {
                lines.Add("  RunPath=нет марша");
                return;
            }

            try
            {
                CurveLoop path = run.GetStairsPath();
                List<Curve> curves = path == null ? new List<Curve>() : path.ToList();
                lines.Add($"  RunPathCurves={curves.Count}");

                double totalPlanLength = 0.0;
                for (int i = 0; i < curves.Count; i++)
                {
                    Curve curve = curves[i];
                    XYZ a = curve.GetEndPoint(0);
                    XYZ b = curve.GetEndPoint(1);
                    double planLength = new XYZ(b.X - a.X, b.Y - a.Y, 0.0).GetLength();
                    totalPlanLength += planLength;
                    lines.Add($"    Curve {i}: {curve.GetType().Name}; P0={FormatXyz(a)}; P1={FormatXyz(b)}; PlanLength={FormatFtMm(planLength)}");
                }

                lines.Add($"  RunPathPlanLengthTotal={FormatFtMm(totalPlanLength)}");
            }
            catch (Exception ex)
            {
                lines.Add($"  RunPath ERROR: {ex.Message}");
            }
        }

        private static void AddDebugElementBlock(List<string> lines, string title, Element elem, bool includeParameters)
        {
            if (lines == null)
                return;

            lines.Add(title);
            if (elem == null)
            {
                lines.Add("  Element=null");
                return;
            }

            lines.Add($"  Id={IDHelper.ElIdValue(elem.Id)}");
            lines.Add($"  Category={GetElementCategoryName(elem)}");
            lines.Add($"  Name={GetElementDisplayName(elem)}");
            lines.Add($"  Type={GetElementTypeDisplayName(elem.Document, elem)}");
            lines.Add($"  UniqueId={elem.UniqueId}");
            lines.Add($"  IfcGUID={GetElementIfcGuid(elem.Document, elem)}");
            lines.Add($"  BoundingBox={FormatBoundingBox(SafeGetBoundingBox(elem))}");
            AddDebugSolidStats(lines, elem, "  Solids");

            if (includeParameters)
                AddDebugParameterDump(lines, elem, "  ");
        }

        private static void AddDebugSolidStats(List<string> lines, Element elem, string prefix)
        {
            if (lines == null)
                return;

            var solids = new List<Solid>();
            AddElementSolids(elem, solids);
            solids = GetValidSolids(solids);
            double volume = solids.Sum(x => x == null ? 0.0 : x.Volume);
            lines.Add($"{prefix}: count={solids.Count}; volume={FormatFt(volume)} ft3");
        }

        private static void AddDebugParameterDump(List<string> lines, Element elem, string indent)
        {
            if (lines == null || elem == null)
                return;

            var values = new List<string>();
            try
            {
                foreach (Parameter parameter in elem.Parameters)
                {
                    string name = parameter?.Definition?.Name ?? "";
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    string value = FormatDebugParameterValue(elem.Document, parameter);
                    if (string.IsNullOrWhiteSpace(value))
                        continue;

                    values.Add($"{name} = {value}");
                }
            }
            catch
            {
            }

            lines.Add($"{indent}Parameters non-empty count={values.Count}");
            foreach (string value in values.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                lines.Add(indent + "  " + value);
        }

        private static string FormatDebugParameterValue(Document doc, Parameter parameter)
        {
            if (parameter == null)
                return "";

            try
            {
                string value = TryGetParameterStringValue(parameter);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch
            {
            }

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.Double:
                        double raw = parameter.AsDouble();
                        return $"{FormatFt(raw)} ft / {FormatMm(raw)} mm";
                    case StorageType.Integer:
                        return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        long value = IDHelper.ElIdValue(id);
                        Element elem = doc == null ? null : doc.GetElement(id);
                        return elem == null ? value.ToString(CultureInfo.InvariantCulture) : $"{value} ({GetElementDisplayName(elem)})";
                    case StorageType.String:
                        return parameter.AsString() ?? "";
                }
            }
            catch
            {
            }

            return "";
        }

        private static void AddDebugReflectedProperties(List<string> lines, object target, string prefix, IEnumerable<string> propertyNames)
        {
            if (lines == null || target == null || propertyNames == null)
                return;

            foreach (string propertyName in propertyNames)
            {
                try
                {
                    PropertyInfo property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                    if (property == null)
                        continue;

                    object value = property.GetValue(target, null);
                    lines.Add($"{prefix}.{propertyName}={FormatDebugObjectValue(value)}");
                }
                catch
                {
                }
            }
        }

        private static string FormatDebugObjectValue(object value)
        {
            if (value == null)
                return "null";

            double doubleValue;
            if (value is double)
            {
                doubleValue = (double)value;
                return $"{FormatFt(doubleValue)} ft / {FormatMm(doubleValue)} mm";
            }

            ElementId elementId = value as ElementId;
            if (elementId != null)
                return IDHelper.ElIdValue(elementId).ToString(CultureInfo.InvariantCulture);

            XYZ xyz = value as XYZ;
            if (xyz != null)
                return FormatXyz(xyz);

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private EvacuationRoutesCheckRequest CreateRouteCheckRequestForRoute(long routeElementId)
        {
            var request = new EvacuationRoutesCheckRequest
            {
                RouteElementId = routeElementId
            };

            DirectShape routeShape = doc?.GetElement(IDHelper.CreateElementId(routeElementId)) as DirectShape;
            if (routeShape == null)
                return request;

            long componentId = GetRouteComponentElementId(routeShape);
            request.ComponentElementId = componentId;
            request.StairElementId = FindOwnerStairIdByComponentId(doc, componentId);
            return request;
        }

        private RouteCheckResult CheckRouteIntersections(EvacuationRoutesCheckRequest request)
        {
            var result = new RouteCheckResult();
            if (request == null || request.RouteElementId <= 0)
            {
                result.ReportText = "Не задан путь эвакуации для проверки.";
                result.HasIntersections = true;
                return result;
            }

            if (doc == null)
            {
                result.ReportText = "Документ Revit недоступен для проверки.";
                result.HasIntersections = true;
                return result;
            }

            DirectShape routeShape = doc.GetElement(IDHelper.CreateElementId(request.RouteElementId)) as DirectShape;
            if (routeShape == null)
            {
                result.ReportText = $"Путь эвакуации ID {request.RouteElementId} не найден.";
                result.HasIntersections = true;
                return result;
            }

            var routeSolids = new List<Solid>();
            AddElementSolids(routeShape, routeSolids);
            routeSolids = GetValidSolids(routeSolids);
            if (routeSolids.Count == 0)
            {
                result.ReportText = $"У пути эвакуации ID {request.RouteElementId} не найдена геометрия для проверки.";
                result.HasIntersections = true;
                return result;
            }

            var reports = new List<RouteIntersectionReportItem>();
            AddRouteIntersectionReport(
                doc,
                routeSolids,
                routeShape,
                string.IsNullOrWhiteSpace(routeShape.Name) ? $"Путь ID {request.RouteElementId}" : routeShape.Name,
                GetRouteCheckExcludedIds(request),
                reports,
                null,
                null,
                IDHelper.CreateElementId(request.ComponentElementId),
                "Элемент");

            List<RouteIntersectionTarget> targets = reports
                .SelectMany(x => x.Targets ?? new List<RouteIntersectionTarget>())
                .OrderBy(x => x.ElementId)
                .ToList();

            result.HasIntersections = targets.Count > 0;
            if (!result.HasIntersections)
            {
                result.ReportText = $"Путь ID {request.RouteElementId}: пересечений не найдено.";
                return result;
            }

            var lines = new List<string> { $"Путь ID {request.RouteElementId}: пересечения ({targets.Count})" };
            foreach (RouteIntersectionTarget target in targets)
                lines.Add("- " + FormatIntersectionTargetShort(target));

            result.ReportText = string.Join(Environment.NewLine, lines);
            return result;
        }

        private HashSet<long> GetRouteCheckExcludedIds(EvacuationRoutesCheckRequest request)
        {
            var ids = new HashSet<long>();
            if (doc == null || request == null)
                return ids;

            Element stair = request.StairElementId > 0 ? doc.GetElement(IDHelper.CreateElementId(request.StairElementId)) : null;
            Stairs stairs = stair as Stairs;
            if (stairs == null)
            {
                MultistoryStairs multistory = stair as MultistoryStairs;
                if (multistory != null)
                    stairs = GetMultistoryStandardStairs(doc, multistory);
            }

            if (stairs != null)
                ids.UnionWith(GetStairAndComponentIds(stairs));

            Element component = request.ComponentElementId > 0 ? doc.GetElement(IDHelper.CreateElementId(request.ComponentElementId)) : null;
            AddElementAndDependentsToExclude(component, ids, 0);

            return ids;
        }

        private RouteEditResult ResizeRouteShape(EvacuationRoutesResizeRequest request)
        {
            var result = new RouteEditResult
            {
                StairElementId = request == null ? 0 : request.StairElementId,
                RouteElementId = request == null ? 0 : request.RouteElementId
            };

            if (request == null || request.RouteElementId <= 0)
            {
                result.Message = "Не задан путь эвакуации для изменения габаритов.";
                return result;
            }

            if (doc == null)
            {
                result.Message = "Документ Revit недоступен для изменения габаритов.";
                return result;
            }

            DirectShape routeShape = doc.GetElement(IDHelper.CreateElementId(request.RouteElementId)) as DirectShape;
            if (routeShape == null)
            {
                result.Message = $"Путь эвакуации ID {request.RouteElementId} не найден.";
                return result;
            }

            RouteDimensionInfo dimensions = GetRouteDimensionInfo(doc, request.RouteElementId, request.ComponentElementId);
            if (dimensions.LengthMm <= 0 || dimensions.WidthMm <= 0 || dimensions.HeightMm <= 0)
            {
                result.Message = $"Не удалось определить текущие габариты пути ID {request.RouteElementId}.";
                return result;
            }

            double newLengthFt = MmToInternal(request.NewLengthMm);
            double newWidthFt = MmToInternal(request.NewWidthMm);
            double newHeightFt = MmToInternal(request.NewHeightMm);
            if (newLengthFt <= 1e-9 || newWidthFt <= 1e-9 || newHeightFt <= 1e-9)
            {
                result.Message = "Новые габариты должны быть положительными.";
                return result;
            }

            double deltaLengthFt = newLengthFt - MmToInternal(dimensions.LengthMm);
            double deltaWidthFt = newWidthFt - MmToInternal(dimensions.WidthMm);
            double deltaHeightFt = newHeightFt - MmToInternal(dimensions.HeightMm);

            var routeSolids = new List<Solid>();
            AddElementSolids(routeShape, routeSolids);
            routeSolids = GetValidSolids(routeSolids);
            if (routeSolids.Count == 0)
            {
                result.Message = $"У пути эвакуации ID {request.RouteElementId} не найдена геометрия для изменения.";
                return result;
            }

            if (!TryGetSolidsLocalExtents(routeSolids, dimensions.XDir, dimensions.YDir, out double minX, out double maxX, out double minY, out double maxY))
            {
                result.Message = $"Не удалось определить крайние грани пути ID {request.RouteElementId}.";
                return result;
            }

            var newShape = new List<GeometryObject>();
            foreach (Solid solid in routeSolids)
            {
                IList<GeometryObject> stretched = TryCreateStretchedSolidGeometry(
                    solid,
                    dimensions.XDir,
                    dimensions.YDir,
                    minX,
                    maxX,
                    minY,
                    maxY,
                    deltaLengthFt,
                    deltaWidthFt,
                    deltaHeightFt,
                    request.LengthDirection,
                    request.WidthDirection);

                if (stretched == null || stretched.Count == 0)
                {
                    result.Message = $"Не удалось перестроить геометрию пути ID {request.RouteElementId}.";
                    return result;
                }

                foreach (GeometryObject geometryObject in stretched)
                    newShape.Add(geometryObject);
            }

            if (newShape.Count == 0)
            {
                result.Message = $"После изменения габаритов у пути ID {request.RouteElementId} не осталось корректной геометрии.";
                return result;
            }

            using (var tx = new Transaction(doc, "KPLN: Изменение габаритов пути эвакуации"))
            {
                tx.Start();
                string appId = routeShape.ApplicationId;
                string appDataId = routeShape.ApplicationDataId;
                string routeName = routeShape.Name;
                routeShape.SetShape(newShape);
                RestoreRouteShapeIdentity(routeShape, appId, appDataId, routeName);
                tx.Commit();
            }

            TryRefreshActiveView();
            try { SelectAndShowElement(uidoc, request.RouteElementId); } catch { }

            RouteCheckResult resizeCheck = CheckRouteIntersections(new EvacuationRoutesCheckRequest
            {
                StairElementId = request.StairElementId,
                ComponentElementId = request.ComponentElementId,
                RouteElementId = request.RouteElementId
            });

            result.CheckReport = resizeCheck.ReportText;
            result.IsFixed = !resizeCheck.HasIntersections;
            FillEditResultDimensions(result, doc, request.RouteElementId, request.ComponentElementId);
            result.Message = $"Габариты пути ID {request.RouteElementId} изменены: длина {request.NewLengthMm:0.#} мм; ширина {request.NewWidthMm:0.#} мм; высота {request.NewHeightMm:0.#} мм.";
            return result;
        }

        private static void FillEditResultDimensions(RouteEditResult result, Document doc, long routeElementId, long componentElementId)
        {
            if (result == null || doc == null || routeElementId <= 0)
                return;

            RouteDimensionInfo dimensions = GetRouteDimensionInfo(doc, routeElementId, componentElementId);
            if (dimensions.LengthMm <= 0 || dimensions.WidthMm <= 0 || dimensions.HeightMm <= 0)
                return;

            result.RouteElementId = routeElementId;
            result.HasDimensions = true;
            result.LengthMm = dimensions.LengthMm;
            result.WidthMm = dimensions.WidthMm;
            result.HeightMm = dimensions.HeightMm;
        }

        private static void RestoreRouteShapeIdentity(DirectShape routeShape, string appId, string appDataId, string routeName)
        {
            if (routeShape == null)
                return;

            try
            {
                if (!string.IsNullOrWhiteSpace(appId))
                    routeShape.ApplicationId = appId;
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(appDataId))
                    routeShape.ApplicationDataId = appDataId;
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(routeName))
                    routeShape.Name = routeName;
            }
            catch
            {
            }
        }

        private static bool TryGetSolidsLocalExtents(IEnumerable<Solid> solids, XYZ xDir, XYZ yDir, out double minX, out double maxX, out double minY, out double maxY)
        {
            xDir = NormalizePlanDir(xDir, XYZ.BasisX);
            yDir = NormalizePlanDir(yDir, XYZ.BasisY);

            minX = double.PositiveInfinity;
            maxX = double.NegativeInfinity;
            minY = double.PositiveInfinity;
            maxY = double.NegativeInfinity;
            bool hasPoint = false;

            foreach (XYZ p in CollectSolidVertices(solids))
            {
                XYZ pxy = new XYZ(p.X, p.Y, 0.0);
                double x = pxy.DotProduct(xDir);
                double y = pxy.DotProduct(yDir);
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
                hasPoint = true;
            }

            return hasPoint && maxX > minX && maxY > minY;
        }

        private static IList<GeometryObject> TryCreateStretchedSolidGeometry(Solid solid, XYZ xDir, XYZ yDir, double minX, double maxX, double minY, double maxY, double deltaLengthFt, double deltaWidthFt, double deltaHeightFt, int lengthDirection, int widthDirection)
        {
            if (solid == null || solid.Volume <= 1e-9)
                return null;

            xDir = NormalizePlanDir(xDir, XYZ.BasisX);
            yDir = NormalizePlanDir(yDir, XYZ.BasisY);
            double tol = MmToInternal(1.0);

            var vertices = CollectSolidVertices(new[] { solid });
            if (vertices.Count == 0)
                return null;

            var topVertexKeys = CollectTopVertexKeys(solid);
            if (topVertexKeys.Count == 0)
            {
                double maxZ = vertices.Max(x => x.Z);
                foreach (XYZ p in vertices.Where(x => Math.Abs(x.Z - maxZ) <= tol))
                    topVertexKeys.Add(GetVertexKey(p, tol));
            }

            var tsb = new TessellatedShapeBuilder();
            tsb.OpenConnectedFaceSet(false);

            foreach (Face face in solid.Faces)
            {
                Mesh mesh = null;
                try { mesh = face.Triangulate(); } catch { }
                if (mesh == null)
                    continue;

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    XYZ a = StretchRouteVertex(triangle.get_Vertex(0), xDir, yDir, minX, maxX, minY, maxY, deltaLengthFt, deltaWidthFt, deltaHeightFt, lengthDirection, widthDirection, topVertexKeys, tol);
                    XYZ b = StretchRouteVertex(triangle.get_Vertex(1), xDir, yDir, minX, maxX, minY, maxY, deltaLengthFt, deltaWidthFt, deltaHeightFt, lengthDirection, widthDirection, topVertexKeys, tol);
                    XYZ c = StretchRouteVertex(triangle.get_Vertex(2), xDir, yDir, minX, maxX, minY, maxY, deltaLengthFt, deltaWidthFt, deltaHeightFt, lengthDirection, widthDirection, topVertexKeys, tol);

                    if (a.DistanceTo(b) <= tol || b.DistanceTo(c) <= tol || c.DistanceTo(a) <= tol)
                        continue;

                    tsb.AddFace(new TessellatedFace(new List<XYZ> { a, b, c }, ElementId.InvalidElementId));
                }
            }

            tsb.CloseConnectedFaceSet();
            tsb.Target = TessellatedShapeBuilderTarget.Solid;
            tsb.Fallback = TessellatedShapeBuilderFallback.Abort;

            try
            {
                tsb.Build();
                TessellatedShapeBuilderResult buildResult = tsb.GetBuildResult();
                return buildResult?.GetGeometricalObjects();
            }
            catch
            {
                return null;
            }
        }

        private static HashSet<string> CollectTopVertexKeys(Solid solid)
        {
            var result = new HashSet<string>();
            double tol = MmToInternal(1.0);

            foreach (Face face in solid.Faces)
            {
                PlanarFace pf = face as PlanarFace;
                if (pf == null || pf.FaceNormal == null || pf.FaceNormal.Z <= 0.05)
                    continue;

                Mesh mesh = null;
                try { mesh = face.Triangulate(); } catch { }
                if (mesh == null)
                    continue;

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    result.Add(GetVertexKey(triangle.get_Vertex(0), tol));
                    result.Add(GetVertexKey(triangle.get_Vertex(1), tol));
                    result.Add(GetVertexKey(triangle.get_Vertex(2), tol));
                }
            }

            return result;
        }

        private static XYZ StretchRouteVertex(
            XYZ p,
            XYZ xDir,
            XYZ yDir,
            double minX,
            double maxX,
            double minY,
            double maxY,
            double deltaLengthFt,
            double deltaWidthFt,
            double deltaHeightFt,
            int lengthDirection,
            int widthDirection,
            HashSet<string> topVertexKeys,
            double tol)
        {
            XYZ result = p;
            XYZ pxy = new XYZ(p.X, p.Y, 0.0);
            double x = pxy.DotProduct(xDir);
            double y = pxy.DotProduct(yDir);

            if (Math.Abs(deltaLengthFt) > 1e-9)
            {
                if (x <= minX + tol)
                    result = result + xDir * GetDirectedMinOffset(deltaLengthFt, lengthDirection);
                else if (x >= maxX - tol)
                    result = result + xDir * GetDirectedMaxOffset(deltaLengthFt, lengthDirection);
            }

            if (Math.Abs(deltaWidthFt) > 1e-9)
            {
                if (y <= minY + tol)
                    result = result + yDir * GetDirectedMinOffset(deltaWidthFt, widthDirection);
                else if (y >= maxY - tol)
                    result = result + yDir * GetDirectedMaxOffset(deltaWidthFt, widthDirection);
            }

            if (Math.Abs(deltaHeightFt) > 1e-9 && topVertexKeys != null && topVertexKeys.Contains(GetVertexKey(p, tol)))
                result = result + XYZ.BasisZ * deltaHeightFt;

            return result;
        }

        private static double GetDirectedMinOffset(double deltaFt, int direction)
        {
            if (direction < 0)
                return -deltaFt;

            if (direction > 0)
                return 0.0;

            return -deltaFt * 0.5;
        }

        private static double GetDirectedMaxOffset(double deltaFt, int direction)
        {
            if (direction < 0)
                return 0.0;

            if (direction > 0)
                return deltaFt;

            return deltaFt * 0.5;
        }

        private static XYZ NormalizePlanDir(XYZ dir, XYZ fallback)
        {
            XYZ result = dir == null ? fallback : new XYZ(dir.X, dir.Y, 0.0);
            if (result == null || result.GetLength() <= 1e-9)
                result = fallback ?? XYZ.BasisX;

            result = new XYZ(result.X, result.Y, 0.0);
            return result.GetLength() <= 1e-9 ? XYZ.BasisX : result.Normalize();
        }

        private static string GetVertexKey(XYZ p, double tol)
        {
            return $"{Math.Round(p.X / tol)}:{Math.Round(p.Y / tol)}:{Math.Round(p.Z / tol)}";
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

        private struct LandingSection2D
        {
            public double X;
            public double MinY;
            public double MaxY;

            public double Width => MaxY - MinY;
        }

        private static string GetElementWorksetName(Document doc, Element elem)
        {
            if (doc == null || elem == null)
                return "";

            if (!doc.IsWorkshared)
                return "Не используется";

            try
            {
                Workset workset = doc.GetWorksetTable().GetWorkset(elem.WorksetId);
                if (workset != null && !string.IsNullOrWhiteSpace(workset.Name))
                    return workset.Name;
            }
            catch
            {
            }

            try
            {
                Parameter parameter = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                string value = parameter == null ? null : parameter.AsValueString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch
            {
            }

            return "Не определён";
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
            debugLog?.Add($"EffectiveUseRunWidth={ShouldUseRunWidth(data, target)}");
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

        private static DirectShape UpsertRouteShape(Document doc, ElementId categoryId, string appId, string appDataId, string name, Solid solid, EvacuationRoutesDialogResult data)
        {
            return UpsertRouteShape(doc, categoryId, appId, appDataId, name, solid == null ? null : new List<Solid> { solid }, data);
        }

        private static DirectShape UpsertRouteShape(Document doc, ElementId categoryId, string appId, string appDataId, string name, IList<Solid> solids, EvacuationRoutesDialogResult data)
        {
            List<Solid> validSolids = GetValidSolids(solids);
            if (validSolids.Count == 0)
                return null;

            DirectShape ds = FindExistingRouteShape(doc, appId, appDataId);

            if (ds == null)
            {
                ds = DirectShape.CreateElement(doc, categoryId);
                ds.ApplicationId = appId;
                ds.ApplicationDataId = appDataId;
            }

            ds.Name = name;

            var geometry = new List<GeometryObject>();
            foreach (Solid solid in validSolids)
                geometry.Add(solid);

            ds.SetShape(geometry);
            TrySetRouteShapeWorkset(ds, data);
            return ds;
        }

        private static List<Solid> GetValidSolids(IEnumerable<Solid> solids)
        {
            var result = new List<Solid>();
            foreach (Solid solid in solids ?? Enumerable.Empty<Solid>())
            {
                if (solid != null && solid.Volume > 1e-9)
                    result.Add(solid);
            }

            return result;
        }

        private static double SumSolidVolumes(IEnumerable<Solid> solids)
        {
            double result = 0.0;
            foreach (Solid solid in solids ?? Enumerable.Empty<Solid>())
            {
                if (solid != null && solid.Volume > 1e-9)
                    result += solid.Volume;
            }

            return result;
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

        private static HashSet<long> GetStairAndComponentIds(Stairs stairs, bool includeAssociatedRailings = true)
        {
            var ids = new HashSet<long>();
            if (stairs == null)
                return ids;

            if (includeAssociatedRailings)
                AddElementAndDependentsToExclude(stairs, ids, 0);
            else
                ids.Add(IDHelper.ElIdValue(stairs.Id));

            try
            {
                var runIds = stairs.GetStairsRuns();
                if (runIds != null)
                {
                    foreach (ElementId id in runIds)
                    {
                        if (includeAssociatedRailings)
                        {
                            Element elem = stairs.Document?.GetElement(id);
                            AddElementAndDependentsToExclude(elem, ids, 0);
                        }
                        else
                        {
                            ids.Add(IDHelper.ElIdValue(id));
                        }
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
                        if (includeAssociatedRailings)
                        {
                            Element elem = stairs.Document?.GetElement(id);
                            AddElementAndDependentsToExclude(elem, ids, 0);
                        }
                        else
                        {
                            ids.Add(IDHelper.ElIdValue(id));
                        }
                    }
                }
            }
            catch
            {
            }

            if (includeAssociatedRailings)
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
            AddRouteIntersectionReport(doc, routeSolid == null ? null : new List<Solid> { routeSolid }, routeShape, routeName, excludedIds, reports, debugLog, target, componentId, componentKind);
        }

        private static void AddRouteIntersectionReport(Document doc, IList<Solid> routeSolids, DirectShape routeShape, string routeName, HashSet<long> excludedIds, List<RouteIntersectionReportItem> reports, RouteDebugLog debugLog, RouteBuildTarget target, ElementId componentId, string componentKind)
        {
            List<Solid> validSolids = GetValidSolids(routeSolids);
            if (doc == null || validSolids.Count == 0 || reports == null)
                return;

            var targets = new List<RouteIntersectionTarget>();
            var seen = new HashSet<string>();

            int hostRawCount = 0;
            int hostReportCount = 0;
            debugLog?.Add($"[INTERSECTION] {routeName}: current stair excluded ids={(excludedIds == null ? 0 : excludedIds.Count)}");

            for (int i = 0; i < validSolids.Count; i++)
            {
                Solid routeSolid = validSolids[i];

                try
                {
                    var candidates = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .WherePasses(new ElementIntersectsSolidFilter(routeSolid))
                        .ToElements();

                    hostRawCount += candidates.Count;
                    debugLog?.Add($"[INTERSECTION] {routeName}: piece {i + 1}/{validSolids.Count}; solidFilter={candidates.Count}");

                    foreach (Element elem in candidates)
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
                    debugLog?.Add($"[INTERSECTION] {routeName}: host check piece {i + 1}/{validSolids.Count} ERROR: {ex.Message}");
                }
            }

            debugLog?.Add($"[INTERSECTION] {routeName}: host pieces={validSolids.Count}; raw={hostRawCount}; reportable={hostReportCount}");

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
            if (cat == null)
                return false;

            return IsReportableIntersectionCategory(cat);
        }

        private static bool IsReportableIntersectionCategory(Category category)
        {
            return IsWallCategory(category)
                || IsFloorCategory(category)
                || IsStairsCategory(category)
                || IsRailingCategory(category);
        }

        private static bool IsWallCategory(Category category)
        {
            return IsBuiltInCategory(category, BuiltInCategory.OST_Walls);
        }

        private static bool IsFloorCategory(Category category)
        {
            return IsBuiltInCategory(category, BuiltInCategory.OST_Floors);
        }

        private static bool IsStairsCategory(Category category)
        {
            return IsBuiltInCategory(category, BuiltInCategory.OST_Stairs);
        }

        private static bool IsRailingCategory(Category category)
        {
            return IsBuiltInCategory(category,
                "OST_Railings",
                "OST_StairsRailing",
                "OST_RailingSystem",
                "OST_RailingRail",
                "OST_RailingTopRail",
                "OST_RailingHandRail",
                "OST_RailingSupport");
        }

        private static bool IsBuiltInCategory(Category category, BuiltInCategory builtInCategory)
        {
            if (category == null)
                return false;

            try
            {
                return IDHelper.ElIdInt(category.Id) == (int)builtInCategory;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsBuiltInCategory(Category category, params string[] builtInCategoryNames)
        {
            if (category == null || builtInCategoryNames == null)
                return false;

            int categoryId = IDHelper.ElIdInt(category.Id);
            foreach (string categoryName in builtInCategoryNames)
            {
                try
                {
                    var bic = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), categoryName);
                    if (categoryId == (int)bic)
                        return true;
                }
                catch
                {
                }
            }

            return false;
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
            double minIntersectionThickness = GetMinIntersectionThicknessFt(elem);

            foreach (Solid elemSolid in solids)
            {
                if (HasMeaningfulSolidIntersection(routeSolid, elemSolid, minIntersectionVolume, minIntersectionThickness))
                    return true;
            }

            return false;
        }

        private static bool HasMeaningfulSolidIntersection(Solid routeSolid, Solid elemSolid, double minIntersectionVolume, double minIntersectionThickness)
        {
            if (routeSolid == null || routeSolid.Volume < 1e-9 || elemSolid == null || elemSolid.Volume < 1e-9)
                return false;

            try
            {
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(routeSolid, elemSolid, BooleanOperationsType.Intersect);
                return intersection != null && intersection.Volume > minIntersectionVolume && HasMinimumIntersectionThickness(intersection, minIntersectionThickness);
            }
            catch
            {
                return false;
            }
        }

        private static double GetMinIntersectionVolumeFt3()
        {
            double oneMmFt = MmToInternal(1.0);
            return oneMmFt * oneMmFt * oneMmFt;
        }

        private static bool HasMinimumIntersectionThickness(Solid intersection, double minThicknessFt)
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
            return effectiveThicknessFt >= minThicknessFt;
        }

        private static double GetMinIntersectionThicknessFt()
        {
            return MmToInternal(5.0);
        }

        private static double GetMinIntersectionThicknessFt(Element elem)
        {
            if (elem != null && IsWallCategory(elem.Category))
                return MmToInternal(30.0);

            return GetMinIntersectionThicknessFt();
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

            if (ShouldUseRunWidth(data, target))
            {
                RunClearWidthInfo clearWidth = GetRunClearWidthInfo(doc, stairs, run, bottomCenter, topCenter, xP, yP, lenPlan, data.ConsiderRailings, debugLog);
                widthFt = clearWidth.WidthFt;
                routeBottomCenter = bottomCenter + yP * clearWidth.CenterOffsetFt;
                routeTopCenter = topCenter + yP * clearWidth.CenterOffsetFt;
                debugLog?.Add($"WidthMode=RunClearWidth; Width={FormatFtMm(widthFt)}; CenterOffset={FormatFtMm(clearWidth.CenterOffsetFt)}");
            }
            else
            {
                double requestedWidthFt = MmToInternal(data.WidthMm);
                double manualCenterOffsetFt;
                RunClearWidthInfo railingCorridor = new RunClearWidthInfo();
                if (data.ConsiderRailings)
                {
                    railingCorridor = GetRunClearWidthInfo(
                        doc,
                        stairs,
                        run,
                        bottomCenter,
                        topCenter,
                        xP,
                        yP,
                        lenPlan,
                        true,
                        debugLog);
                }

                if (data.ConsiderRailings
                    && railingCorridor.HasRailingBoundary
                    && railingCorridor.WidthFt > 1e-9)
                {
                    widthFt = railingCorridor.WidthFt;
                    manualCenterOffsetFt = railingCorridor.CenterOffsetFt;
                    debugLog?.Add(
                        $"ManualWidth replaced by railing corridor; requested={FormatFtMm(requestedWidthFt)}; " +
                        $"left={railingCorridor.HasLeftRailingBoundary}; right={railingCorridor.HasRightRailingBoundary}; " +
                        $"corridor={FormatFtMm(railingCorridor.WidthFt)}; result={FormatFtMm(widthFt)}");
                }
                else
                {
                    widthFt = requestedWidthFt;
                    manualCenterOffsetFt = GetRunManualWidthCenterOffset(run, bottomCenter, topCenter, xP, yP, lenPlan, debugLog);
                }

                routeBottomCenter = bottomCenter + yP * manualCenterOffsetFt;
                routeTopCenter = topCenter + yP * manualCenterOffsetFt;
                debugLog?.Add($"WidthMode=Manual; Width={FormatFtMm(widthFt)}; CenterOffset={FormatFtMm(manualCenterOffsetFt)}");
            }

            if (data.RoundRunWidthDownTo5Mm)
            {
                double sourceWidthFt = widthFt;
                widthFt = RoundWidthDownTo5Mm(widthFt);
                debugLog?.Add($"WidthRoundDownTo5Mm: Source={FormatFtMm(sourceWidthFt)}; Result={FormatFtMm(widthFt)}");
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

            string routeName = CreateRouteName(target, run.Id, isLanding: false);
            string appDataId = CreateRouteAppDataId(target, run.Id);
            DirectShape routeShape = UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", appDataId, routeName, solid, data);
            AddRouteIntersectionReport(doc, solid, routeShape, routeName, GetStairAndComponentIds(stairs), intersectionReports, debugLog, target, run.Id, "Марш");
            return true;
        }

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

            var candidates = new List<(RunRouteBodyInfo run, EndFace face, XYZ faceCenter, double dist, string endName)>();
            foreach (var ri in runInfos)
            {
                double dTopPlan = GetPlanDistanceToBoundingBox(ri.TopEnd.Center, bbL);
                double dBotPlan = GetPlanDistanceToBoundingBox(ri.BottomEnd.Center, bbL);
                double dTopZ = GetDistanceToRange(ri.TopEnd.MinZBottom, bbL.Min.Z, bbL.Max.Z);
                double dBotZ = GetDistanceToRange(ri.BottomEnd.MinZBottom, bbL.Min.Z, bbL.Max.Z);
                double dTop = dTopPlan + dTopZ;
                double dBot = dBotPlan + dBotZ;

                bool useTop = dTop <= dBot;
                EndFace f = useTop ? ri.TopEnd : ri.BottomEnd;

                debugLog?.Add($"Candidate run={IDHelper.ElIdValue(ri.RunId)} dTop={FormatFtMm(dTop)} (plan={FormatFtMm(dTopPlan)} z={FormatFtMm(dTopZ)}) dBottom={FormatFtMm(dBot)} (plan={FormatFtMm(dBotPlan)} z={FormatFtMm(dBotZ)}) selected={(useTop ? "Top" : "Bottom")} selectedCenter={FormatXyz(f.Center)} selectedMinZ={FormatFtMm(f.MinZBottom)} width={FormatFtMm(ri.WidthFt)}");

                candidates.Add((ri, f, f.Center, useTop ? dTop : dBot, useTop ? "Top" : "Bottom"));
            }

            if (candidates.Count < 2)
                return false;

            var two = candidates.OrderBy(x => x.dist).Take(2).ToList();
            var A = two[0];
            var B = two[1];

            debugLog?.Add($"SelectedA run={IDHelper.ElIdValue(A.run.RunId)} end={A.endName} dist={FormatFtMm(A.dist)} center={FormatXyz(A.faceCenter)}");
            debugLog?.Add($"SelectedB run={IDHelper.ElIdValue(B.run.RunId)} end={B.endName} dist={FormatFtMm(B.dist)} center={FormatXyz(B.faceCenter)}");

            XYZ xDir = new XYZ(B.faceCenter.X - A.faceCenter.X, B.faceCenter.Y - A.faceCenter.Y, 0);
            if (xDir.GetLength() < 1e-9)
                xDir = new XYZ(A.run.XDirPlan.X, A.run.XDirPlan.Y, 0);
            if (xDir.GetLength() < 1e-9)
                return false;
            xDir = xDir.Normalize();

            debugLog?.Add($"LandingXDir={FormatXyz(xDir)}");

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


            double manualDepthFt = MmToInternal(data.WidthMm);
            bool useRunWidthsForLanding = UseExperimentalLandingNarrowSection && ShouldUseRunWidth(data, target);
            double depthAFt = useRunWidthsForLanding ? A.run.WidthFt : manualDepthFt;
            double depthBFt = useRunWidthsForLanding ? B.run.WidthFt : manualDepthFt;
            double depthFt = Math.Max(depthAFt, depthBFt);

            if (depthAFt <= 1e-9 || depthBFt <= 1e-9 || depthFt <= 1e-9)
                return false;

            double landingCY = new XYZ(landingCenter.X, landingCenter.Y, 0).DotProduct(yDir);
            double tol = MmToInternal(2.0);

            debugLog?.Add(
                $"LandingWidthMode={(useRunWidthsForLanding ? "ExperimentalRunWidths" : "ManualWidth")}; " +
                $"runA={FormatFtMm(depthAFt)}; runB={FormatFtMm(depthBFt)}; commonDepth={FormatFtMm(depthFt)}");

            double minY, maxY;
            string yDecision;

            if (landingCY > spanMaxY + tol)
            {
                minY = spanMaxY;
                maxY = spanMaxY + depthFt;
                yDecision = "landingCY > spanMaxY + tol";
            }
            else if (landingCY < spanMinY - tol)
            {
                maxY = spanMinY;
                minY = spanMinY - depthFt;
                yDecision = "landingCY < spanMinY - tol";
            }
            else
            {
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

            double baseZ = Math.Min(A.face.MinZBottom, B.face.MinZBottom);
            double h = new[] { heightFt, A.run.HeightFt, B.run.HeightFt }
                .Where(value => value > 1e-9)
                .DefaultIfEmpty(heightFt)
                .Min();
            if (h <= 1e-9) return false;

            debugLog?.Add($"LandingBaseZ={FormatFtMm(baseZ)} height={FormatFtMm(h)}");

            XYZ candidateXDir = xDir;
            XYZ candidateYDir = yDir;
            List<LocalRect2D> candidateRects = BuildLandingCandidateRects(
                A.run, A.face, A.faceCenter,
                B.run, B.face, B.faceCenter,
                xDir, yDir, minX, maxX, minY, maxY, depthFt, depthAFt, depthBFt, tol,
                out candidateXDir, out candidateYDir, debugLog);

            List<Solid> routeSolids = null;
            XYZ sectionXDir;
            XYZ sectionYDir;
            int sectionLandingDirection;
            bool hasParallelRunAxes = TryGetParallelLandingSectionAxes(
                A.run,
                B.run,
                A.faceCenter,
                B.faceCenter,
                landingCenter,
                out sectionXDir,
                out sectionYDir,
                out sectionLandingDirection);

            if (hasParallelRunAxes
                && TryResolveLandingDirectionFromRunEnds(
                    A.run,
                    A.endName,
                    B.run,
                    B.endName,
                    sectionYDir,
                    out int runEndLandingDirection))
            {
                sectionLandingDirection = runEndLandingDirection;
            }

            double sectionScanMinX = 0.0;
            double sectionScanMaxX = 0.0;
            double sectionInnerAnchorY = 0.0;
            double sectionRunAMinX = 0.0;
            double sectionRunAMaxX = 0.0;
            double sectionRunAFaceY = 0.0;
            double sectionRunBMinX = 0.0;
            double sectionRunBMaxX = 0.0;
            double sectionRunBFaceY = 0.0;
            if (hasParallelRunAxes)
            {
                A.face.GetSpanOnDir(sectionXDir, out double aSectionMinX, out double aSectionMaxX);
                B.face.GetSpanOnDir(sectionXDir, out double bSectionMinX, out double bSectionMaxX);
                sectionRunAMinX = aSectionMinX;
                sectionRunAMaxX = aSectionMaxX;
                sectionRunBMinX = bSectionMinX;
                sectionRunBMaxX = bSectionMaxX;
                sectionScanMinX = Math.Min(aSectionMinX, bSectionMinX);
                sectionScanMaxX = Math.Max(aSectionMaxX, bSectionMaxX);
                sectionRunAFaceY = new XYZ(A.faceCenter.X, A.faceCenter.Y, 0.0)
                    .DotProduct(sectionYDir);
                sectionRunBFaceY = new XYZ(B.faceCenter.X, B.faceCenter.Y, 0.0)
                    .DotProduct(sectionYDir);
                XYZ averageFaceCenter = (A.faceCenter + B.faceCenter) * 0.5;
                sectionInnerAnchorY = new XYZ(averageFaceCenter.X, averageFaceCenter.Y, 0.0)
                    .DotProduct(sectionYDir);
            }

            bool usedComplexFootprint = UseExperimentalLandingNarrowSection
                && hasParallelRunAxes
                && TryCreateComplexLandingRouteSolidsFromFootprint(
                    landing,
                    sectionXDir,
                    sectionYDir,
                    sectionScanMinX,
                    sectionScanMaxX,
                    sectionLandingDirection,
                    sectionInnerAnchorY,
                    depthFt,
                    baseZ,
                    h,
                    debugLog,
                    out routeSolids);

            bool usedSectionLimitedFootprint = !usedComplexFootprint
                && UseExperimentalLandingNarrowSection
                && hasParallelRunAxes
                && sectionScanMaxX - sectionScanMinX > tol
                && TryCreateLandingRouteSolidsBySectionLimit(
                    landing,
                    sectionXDir,
                    sectionYDir,
                    sectionLandingDirection,
                    sectionScanMinX,
                    sectionScanMaxX,
                    sectionInnerAnchorY,
                    depthFt,
                    baseZ,
                    h,
                    debugLog,
                    out routeSolids);

            if (!usedComplexFootprint
                && !usedSectionLimitedFootprint
                && !TryCreateLandingRouteSolidsFromFootprint(landing, candidateXDir, candidateYDir, candidateRects, baseZ, h, debugLog, out routeSolids))
            {
                debugLog?.Add("LandingFootprintClip=fallback rectangle");
                XYZ P1 = xDir * minX + yDir * minY + XYZ.BasisZ * baseZ;
                XYZ P2 = xDir * maxX + yDir * minY + XYZ.BasisZ * baseZ;
                XYZ P3 = xDir * maxX + yDir * maxY + XYZ.BasisZ * baseZ;
                XYZ P4 = xDir * minX + yDir * maxY + XYZ.BasisZ * baseZ;
                XYZ up = XYZ.BasisZ * h;
                Solid solid = BuildPrismFrom8Points(P1, P2, P3, P4, P1 + up, P2 + up, P3 + up, P4 + up);
                if (solid == null || solid.Volume < 1e-9)
                    return false;

                routeSolids = new List<Solid> { solid };
            }

            if (hasParallelRunAxes && routeSolids != null && routeSolids.Count > 0)
            {
                int addedRunConnectors = AddLandingRunFaceConnectors(
                    landing,
                    sectionXDir,
                    sectionYDir,
                    sectionScanMinX,
                    sectionScanMaxX,
                    sectionLandingDirection,
                    sectionInnerAnchorY,
                    sectionRunAMinX,
                    sectionRunAMaxX,
                    sectionRunAFaceY,
                    sectionRunBMinX,
                    sectionRunBMaxX,
                    sectionRunBFaceY,
                    baseZ,
                    h,
                    routeSolids,
                    debugLog);
                if (addedRunConnectors > 0)
                    routeSolids = UnionConnectedLandingSolids(routeSolids, debugLog);
            }

            string routeName = CreateRouteName(target, landing.Id, isLanding: true);
            string appDataId = CreateRouteAppDataId(target, landing.Id);
            DirectShape routeShape = UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", appDataId, routeName, routeSolids, data);
            AddRouteIntersectionReport(doc, routeSolids, routeShape, routeName, GetStairAndComponentIds(stairs), intersectionReports, debugLog, target, landing.Id, "Площадка");
            return true;
        }

        private static double GetPlanDistanceToBoundingBox(XYZ point, BoundingBoxXYZ box)
        {
            if (point == null || box == null)
                return double.PositiveInfinity;

            double dx = point.X < box.Min.X
                ? box.Min.X - point.X
                : point.X > box.Max.X ? point.X - box.Max.X : 0.0;
            double dy = point.Y < box.Min.Y
                ? box.Min.Y - point.Y
                : point.Y > box.Max.Y ? point.Y - box.Max.Y : 0.0;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static double GetDistanceToRange(double value, double min, double max)
        {
            if (value < min) return min - value;
            if (value > max) return value - max;
            return 0.0;
        }

        private static List<LocalRect2D> BuildLandingCandidateRects(
            RunRouteBodyInfo runA,
            EndFace faceA,
            XYZ faceCenterA,
            RunRouteBodyInfo runB,
            EndFace faceB,
            XYZ faceCenterB,
            XYZ fallbackXDir,
            XYZ fallbackYDir,
            double fallbackMinX,
            double fallbackMaxX,
            double fallbackMinY,
            double fallbackMaxY,
            double depthFt,
            double depthAFt,
            double depthBFt,
            double tolFt,
            out XYZ candidateXDir,
            out XYZ candidateYDir,
            RouteDebugLog debugLog)
        {
            candidateXDir = fallbackXDir;
            candidateYDir = fallbackYDir;

            var rectangle = new List<LocalRect2D>
            {
                new LocalRect2D(fallbackMinX, fallbackMaxX, fallbackMinY, fallbackMaxY, "rectangle")
            };

            double lThreshold = Math.Max(MmToInternal(100.0), Math.Min(depthFt * 0.35, MmToInternal(300.0)));
            double dx;
            double dy;
            string axisSource;

            if (!TrySelectLandingCandidateAxes(runA, runB, faceCenterA, faceCenterB, fallbackXDir, fallbackYDir, lThreshold, out candidateXDir, out candidateYDir, out dx, out dy, out axisSource))
            {
                debugLog?.Add($"LandingCandidate=rectangle; reason=aligned; threshold={FormatFtMm(lThreshold)}");
                return rectangle;
            }

            double aMinX, aMaxX, aMinY, aMaxY;
            double bMinX, bMaxX, bMinY, bMaxY;
            faceA.GetSpanOnDir(candidateXDir, out aMinX, out aMaxX);
            faceA.GetSpanOnDir(candidateYDir, out aMinY, out aMaxY);
            faceB.GetSpanOnDir(candidateXDir, out bMinX, out bMaxX);
            faceB.GetSpanOnDir(candidateYDir, out bMinY, out bMaxY);

            XYZ aLocal = ProjectLocal2D(faceCenterA, candidateXDir, candidateYDir);
            XYZ bLocal = ProjectLocal2D(faceCenterB, candidateXDir, candidateYDir);

            double aLegMinX = aMinX;
            double aLegMaxX = aMaxX;
            double bLegMinX = bMinX;
            double bLegMaxX = bMaxX;
            double aBandMinY = aMinY;
            double aBandMaxY = aMaxY;
            double bBandMinY = bMinY;
            double bBandMaxY = bMaxY;

            EnsureSpanAtLeast(ref aLegMinX, ref aLegMaxX, aLocal.X, depthAFt);
            EnsureSpanAtLeast(ref bLegMinX, ref bLegMaxX, bLocal.X, depthBFt);
            EnsureSpanAtLeast(ref aBandMinY, ref aBandMaxY, aLocal.Y, depthAFt);
            EnsureSpanAtLeast(ref bBandMinY, ref bBandMaxY, bLocal.Y, depthBFt);

            double pad = MmToInternal(5.0);
            var rects = new List<LocalRect2D>
            {
                ExpandLocalRect(new LocalRect2D(Math.Min(aMinX, bMinX), Math.Max(aMaxX, bMaxX), aBandMinY, aBandMaxY, "L-main-A", "L_A_to_B"), pad),
                ExpandLocalRect(new LocalRect2D(bLegMinX, bLegMaxX, Math.Min(aBandMinY, bBandMinY), Math.Max(aBandMaxY, bBandMaxY), "L-pocket-B", "L_A_to_B"), pad),
                ExpandLocalRect(new LocalRect2D(Math.Min(aMinX, bMinX), Math.Max(aMaxX, bMaxX), bBandMinY, bBandMaxY, "L-main-B", "L_B_to_A"), pad),
                ExpandLocalRect(new LocalRect2D(aLegMinX, aLegMaxX, Math.Min(aBandMinY, bBandMinY), Math.Max(aBandMaxY, bBandMaxY), "L-pocket-A", "L_B_to_A"), pad)
            };

            rects = rects
                .Where(x => x.MaxX - x.MinX > tolFt && x.MaxY - x.MinY > tolFt)
                .ToList();

            if (rects.Count == 0)
            {
                candidateXDir = fallbackXDir;
                candidateYDir = fallbackYDir;
                debugLog?.Add("LandingCandidate=rectangle; reason=empty L rects");
                return rectangle;
            }

            debugLog?.Add($"LandingCandidate=L; axes={axisSource}; dx={FormatFtMm(dx)}; dy={FormatFtMm(dy)}; threshold={FormatFtMm(lThreshold)}; variants={rects.Select(x => x.Group).Distinct().Count()}; rects={rects.Count}");
            return rects;
        }

        private static bool TryGetParallelLandingSectionAxes(
            RunRouteBodyInfo runA,
            RunRouteBodyInfo runB,
            XYZ faceCenterA,
            XYZ faceCenterB,
            XYZ landingCenter,
            out XYZ xDir,
            out XYZ yDir,
            out int landingDirection)
        {
            xDir = null;
            yDir = null;
            landingDirection = 1;
            if (runA == null || runB == null || runA.XDirPlan == null || runB.XDirPlan == null)
                return false;

            XYZ aDir = NormalizePlanDir(runA.XDirPlan, XYZ.BasisX);
            XYZ bDir = NormalizePlanDir(runB.XDirPlan, XYZ.BasisX);
            double parallelDot = Math.Abs(aDir.DotProduct(bDir));
            if (parallelDot < Math.Cos(Math.PI / 18.0))
                return false;

            yDir = aDir;
            xDir = XYZ.BasisZ.CrossProduct(yDir);
            if (xDir == null || xDir.GetLength() <= 1e-9)
                return false;

            xDir = xDir.Normalize();

            XYZ averageFaceCenter = (faceCenterA + faceCenterB) * 0.5;
            double faceY = new XYZ(averageFaceCenter.X, averageFaceCenter.Y, 0.0).DotProduct(yDir);
            double landingY = new XYZ(landingCenter.X, landingCenter.Y, 0.0).DotProduct(yDir);
            landingDirection = landingY >= faceY ? 1 : -1;
            return true;
        }

        private static bool TryResolveLandingDirectionFromRunEnds(
            RunRouteBodyInfo runA,
            string endNameA,
            RunRouteBodyInfo runB,
            string endNameB,
            XYZ sectionYDir,
            out int landingDirection)
        {
            landingDirection = 1;
            if (runA == null || runB == null || sectionYDir == null || sectionYDir.GetLength() <= 1e-9)
                return false;

            XYZ aDir = NormalizePlanDir(runA.XDirPlan, XYZ.BasisX);
            XYZ bDir = NormalizePlanDir(runB.XDirPlan, XYZ.BasisX);
            XYZ aIntoLanding = string.Equals(endNameA, "Top", StringComparison.OrdinalIgnoreCase) ? aDir : aDir * -1.0;
            XYZ bIntoLanding = string.Equals(endNameB, "Top", StringComparison.OrdinalIgnoreCase) ? bDir : bDir * -1.0;
            XYZ yDir = NormalizePlanDir(sectionYDir, XYZ.BasisX);

            double score = aIntoLanding.DotProduct(yDir) + bIntoLanding.DotProduct(yDir);
            if (Math.Abs(score) <= 0.5)
                return false;

            landingDirection = score >= 0.0 ? 1 : -1;
            return true;
        }

        private static bool TryCreateComplexLandingRouteSolidsFromFootprint(
            StairsLanding landing,
            XYZ xDir,
            XYZ yDir,
            double scanMinX,
            double scanMaxX,
            int landingDirection,
            double innerAnchorY,
            double targetDepthFt,
            double baseZ,
            double heightFt,
            RouteDebugLog debugLog,
            out List<Solid> routeSolids)
        {
            routeSolids = new List<Solid>();
            if (landing == null || targetDepthFt <= 1e-9 || heightFt <= 1e-9)
                return false;

            List<List<XYZ>> footprintLoops;
            if (!TryGetLandingTopFootprintLoops2D(landing, xDir, yDir, debugLog, out footprintLoops))
                return false;

            double tolFt = MmToInternal(1.0);
            double contactToleranceFt = MmToInternal(100.0);
            List<XYZ> selectedLoop = null;
            double selectedArea = 0.0;

            foreach (List<XYZ> rawLoop in footprintLoops)
            {
                List<XYZ> loop = CleanPolygon2D(rawLoop, tolFt);
                if (loop.Count < 6 || !IsConcavePolygon2D(loop, tolFt))
                    continue;

                double minX = loop.Min(point => point.X);
                double maxX = loop.Max(point => point.X);
                double minY = loop.Min(point => point.Y);
                double maxY = loop.Max(point => point.Y);
                double overlapX = Math.Min(maxX, scanMaxX) - Math.Max(minX, scanMinX);
                double anchorDistance = GetDistanceToRange(innerAnchorY, minY, maxY);
                if (overlapX <= tolFt || anchorDistance > contactToleranceFt)
                    continue;

                double area = Math.Abs(GetSignedPolygonArea2D(loop));
                if (area <= selectedArea)
                    continue;

                selectedLoop = loop;
                selectedArea = area;
            }

            if (selectedLoop == null)
                return false;

            double bandMinY = landingDirection >= 0
                ? innerAnchorY
                : innerAnchorY - targetDepthFt;
            double bandMaxY = landingDirection >= 0
                ? innerAnchorY + targetDepthFt
                : innerAnchorY;
            List<XYZ> clippedLoop = ClipPolygonToRect2D(
                selectedLoop,
                scanMinX,
                scanMaxX,
                bandMinY,
                bandMaxY,
                tolFt);
            clippedLoop = CleanPolygon2D(clippedLoop, tolFt);
            if (clippedLoop.Count < 3 || Math.Abs(GetSignedPolygonArea2D(clippedLoop)) <= tolFt * tolFt)
                return false;

            Solid complexSolid = TryCreateVerticalExtrusionFromLocalPolygon2D(
                clippedLoop,
                xDir,
                yDir,
                baseZ,
                heightFt,
                tolFt);
            if (complexSolid == null || complexSolid.Volume <= 1e-9)
                return false;

            routeSolids.Add(complexSolid);
            debugLog?.Add($"LandingComplexFootprint=ok; vertices={selectedLoop.Count}->{clippedLoop.Count}; area={FormatFt(selectedArea)}; depth={FormatFtMm(targetDepthFt)}; x={FormatFt(scanMinX)}..{FormatFt(scanMaxX)}; bandY={FormatFt(bandMinY)}..{FormatFt(bandMaxY)}; shape=concave-clipped");
            return routeSolids.Count > 0;
        }

        private static int AddLandingRunFaceConnectors(
            StairsLanding landing,
            XYZ xDir,
            XYZ yDir,
            double scanMinX,
            double scanMaxX,
            int landingDirection,
            double innerAnchorY,
            double runAMinX,
            double runAMaxX,
            double runAFaceY,
            double runBMinX,
            double runBMaxX,
            double runBFaceY,
            double baseZ,
            double heightFt,
            List<Solid> routeSolids,
            RouteDebugLog debugLog)
        {
            if (landing == null || routeSolids == null || xDir == null || yDir == null)
                return 0;

            List<List<XYZ>> footprintLoops;
            if (!TryGetLandingTopFootprintLoops2D(landing, xDir, yDir, debugLog, out footprintLoops))
                return 0;

            double tolFt = MmToInternal(1.0);
            double contactToleranceFt = MmToInternal(300.0);
            List<XYZ> selectedLoop = footprintLoops
                .Select(loop => CleanPolygon2D(loop, tolFt))
                .Where(loop => loop.Count >= 3)
                .Where(loop =>
                {
                    double minX = loop.Min(point => point.X);
                    double maxX = loop.Max(point => point.X);
                    double minY = loop.Min(point => point.Y);
                    double maxY = loop.Max(point => point.Y);
                    double overlapX = Math.Min(maxX, scanMaxX) - Math.Max(minX, scanMinX);
                    return overlapX > tolFt
                        && GetDistanceToRange(innerAnchorY, minY, maxY) <= contactToleranceFt;
                })
                .OrderByDescending(loop => Math.Abs(GetSignedPolygonArea2D(loop)))
                .FirstOrDefault();
            if (selectedLoop == null)
                return 0;

            int connectorCount = 0;
            if (TryAddLandingRunFaceConnector(
                routeSolids,
                selectedLoop,
                xDir,
                yDir,
                runAMinX,
                runAMaxX,
                runAFaceY,
                landingDirection,
                innerAnchorY,
                baseZ,
                heightFt,
                tolFt,
                debugLog,
                "A"))
            {
                connectorCount++;
            }

            if (TryAddLandingRunFaceConnector(
                routeSolids,
                selectedLoop,
                xDir,
                yDir,
                runBMinX,
                runBMaxX,
                runBFaceY,
                landingDirection,
                innerAnchorY,
                baseZ,
                heightFt,
                tolFt,
                debugLog,
                "B"))
            {
                connectorCount++;
            }

            debugLog?.Add($"LandingRunConnectors final count={connectorCount}; anchorY={FormatFt(innerAnchorY)}");
            return connectorCount;
        }

        private static bool TryAddLandingRunFaceConnector(
            List<Solid> routeSolids,
            List<XYZ> landingLoop,
            XYZ xDir,
            XYZ yDir,
            double runMinX,
            double runMaxX,
            double runFaceY,
            int landingDirection,
            double innerAnchorY,
            double baseZ,
            double heightFt,
            double tolFt,
            RouteDebugLog debugLog,
            string runLabel)
        {
            if (routeSolids == null
                || landingLoop == null
                || landingLoop.Count < 3
                || runMaxX - runMinX <= tolFt)
            {
                return false;
            }

            double connectorMinX = 0.0;
            double connectorMaxX = 0.0;
            double sectionY = innerAnchorY;
            bool hasConnectorRange = false;
            double directionSign = landingDirection >= 0 ? 1.0 : -1.0;
            double[] sectionOffsetsMm = { 5.0, 20.0, 50.0, 100.0, -5.0, -20.0 };
            foreach (double offsetMm in sectionOffsetsMm)
            {
                double candidateY = innerAnchorY + directionSign * MmToInternal(offsetMm);
                if (!TryGetLandingXRangeAtY(
                    landingLoop,
                    candidateY,
                    runMinX,
                    runMaxX,
                    tolFt,
                    out connectorMinX,
                    out connectorMaxX))
                {
                    continue;
                }

                sectionY = candidateY;
                hasConnectorRange = true;
                break;
            }

            if (!hasConnectorRange)
                return false;

            double overlapFt = MmToInternal(10.0);
            double connectorMinY = Math.Min(runFaceY, innerAnchorY) - overlapFt;
            double connectorMaxY = Math.Max(runFaceY, innerAnchorY) + overlapFt;

            if (connectorMaxY - connectorMinY <= tolFt)
                return false;

            var connectorPolygon = new List<XYZ>
            {
                new XYZ(connectorMinX, connectorMinY, 0.0),
                new XYZ(connectorMaxX, connectorMinY, 0.0),
                new XYZ(connectorMaxX, connectorMaxY, 0.0),
                new XYZ(connectorMinX, connectorMaxY, 0.0)
            };
            Solid connectorSolid = TryCreateVerticalExtrusionFromLocalPolygon2D(
                connectorPolygon,
                xDir,
                yDir,
                baseZ,
                heightFt,
                tolFt);
            if (connectorSolid == null || connectorSolid.Volume <= 1e-9)
                return false;

            routeSolids.Add(connectorSolid);
            debugLog?.Add(
                $"LandingRunConnector {runLabel}; runX={FormatFt(runMinX)}..{FormatFt(runMaxX)}; " +
                $"connectorX={FormatFt(connectorMinX)}..{FormatFt(connectorMaxX)}; " +
                $"faceY={FormatFt(runFaceY)}; anchorY={FormatFt(innerAnchorY)}; sectionY={FormatFt(sectionY)}; " +
                $"y={FormatFt(connectorMinY)}..{FormatFt(connectorMaxY)}");
            return true;
        }

        private static bool TryGetLandingXRangeAtY(
            List<XYZ> polygon,
            double y,
            double requestedMinX,
            double requestedMaxX,
            double tolFt,
            out double resultMinX,
            out double resultMaxX)
        {
            resultMinX = 0.0;
            resultMaxX = 0.0;
            if (polygon == null || polygon.Count < 3 || requestedMaxX - requestedMinX <= tolFt)
                return false;

            var intersections = new List<double>();
            for (int i = 0; i < polygon.Count; i++)
            {
                XYZ a = polygon[i];
                XYZ b = polygon[(i + 1) % polygon.Count];
                if (a == null || b == null || Math.Abs(b.Y - a.Y) <= tolFt)
                    continue;

                double edgeMinY = Math.Min(a.Y, b.Y);
                double edgeMaxY = Math.Max(a.Y, b.Y);
                if (y < edgeMinY || y >= edgeMaxY)
                    continue;

                double t = (y - a.Y) / (b.Y - a.Y);
                intersections.Add(a.X + (b.X - a.X) * t);
            }

            intersections = intersections
                .OrderBy(x => x)
                .Aggregate(new List<double>(), (result, value) =>
                {
                    if (result.Count == 0 || Math.Abs(value - result[result.Count - 1]) > tolFt)
                        result.Add(value);
                    return result;
                });

            double bestOverlap = 0.0;
            for (int i = 0; i + 1 < intersections.Count; i += 2)
            {
                double minX = Math.Max(requestedMinX, intersections[i]);
                double maxX = Math.Min(requestedMaxX, intersections[i + 1]);
                double overlap = maxX - minX;
                if (overlap <= bestOverlap || overlap <= tolFt)
                    continue;

                bestOverlap = overlap;
                resultMinX = minX;
                resultMaxX = maxX;
            }

            return bestOverlap > tolFt;
        }

        private static bool IsConcavePolygon2D(List<XYZ> polygon, double tolFt)
        {
            List<XYZ> clean = CleanPolygon2D(polygon, tolFt);
            if (clean.Count < 4)
                return false;

            int turnSign = 0;
            for (int i = 0; i < clean.Count; i++)
            {
                XYZ a = clean[i];
                XYZ b = clean[(i + 1) % clean.Count];
                XYZ c = clean[(i + 2) % clean.Count];
                double cross = (b.X - a.X) * (c.Y - b.Y) - (b.Y - a.Y) * (c.X - b.X);
                if (Math.Abs(cross) <= tolFt * tolFt)
                    continue;

                int currentSign = cross > 0.0 ? 1 : -1;
                if (turnSign == 0)
                    turnSign = currentSign;
                else if (turnSign != currentSign)
                    return true;
            }

            return false;
        }

        private static bool TryCreateLandingRouteSolidsBySectionLimit(
            StairsLanding landing,
            XYZ xDir,
            XYZ yDir,
            int landingDirection,
            double scanMinX,
            double scanMaxX,
            double innerAnchorY,
            double targetWidthFt,
            double baseZ,
            double heightFt,
            RouteDebugLog debugLog,
            out List<Solid> routeSolids)
        {
            routeSolids = new List<Solid>();
            if (landing == null || targetWidthFt <= 1e-9 || heightFt <= 1e-9)
                return false;

            List<List<XYZ>> footprintLoops;
            bool hasFootprint = TryGetLandingTopFootprintLoops2D(
                landing,
                xDir,
                yDir,
                debugLog,
                out footprintLoops);
            if (!hasFootprint)
                footprintLoops = new List<List<XYZ>>();

            List<double> criticalX = footprintLoops
                .SelectMany(loop => loop ?? new List<XYZ>())
                .Where(point => point != null)
                .Select(point => point.X)
                .Where(value => value > scanMinX && value < scanMaxX)
                .Concat(new[] { scanMinX, scanMaxX })
                .OrderBy(value => value)
                .Aggregate(new List<double>(), (result, value) =>
                {
                    double mergeToleranceFt = MmToInternal(1.0);
                    if (result.Count == 0 || Math.Abs(value - result[result.Count - 1]) > mergeToleranceFt)
                        result.Add(value);
                    return result;
                });

            if (criticalX.Count < 2)
                return false;

            double tolFt = MmToInternal(1.0);
            double contactToleranceFt = MmToInternal(100.0);
            var landingCells = new List<(double leftX, double rightX, LandingSection2D section)>();
            double positiveCapacity = 0.0;
            double negativeCapacity = 0.0;

            for (int i = 0; i + 1 < criticalX.Count; i++)
            {
                double leftX = criticalX[i];
                double rightX = criticalX[i + 1];
                double cellWidth = rightX - leftX;
                if (cellWidth <= tolFt)
                    continue;

                double sampleX = (leftX + rightX) * 0.5;
                LandingSection2D section;
                if (!TryGetLandingSectionClosestToY(
                        footprintLoops,
                        sampleX,
                        innerAnchorY,
                        tolFt,
                        out section))
                {
                    continue;
                }

                double contactDistance = GetDistanceToRange(innerAnchorY, section.MinY, section.MaxY);
                if (contactDistance > contactToleranceFt)
                {
                    debugLog?.Add($"LandingCell skipped no-run-contact x={FormatFt(sampleX)} distance={FormatFtMm(contactDistance)}");
                    continue;
                }

                landingCells.Add((leftX, rightX, section));
                positiveCapacity += cellWidth * Math.Max(0.0, section.MaxY - innerAnchorY);
                negativeCapacity += cellWidth * Math.Max(0.0, innerAnchorY - section.MinY);
            }

            if (landingCells.Count == 0)
            {
                var fallbackSection = new LandingSection2D
                {
                    X = (scanMinX + scanMaxX) * 0.5,
                    MinY = landingDirection >= 0 ? innerAnchorY : innerAnchorY - targetWidthFt,
                    MaxY = landingDirection >= 0 ? innerAnchorY + targetWidthFt : innerAnchorY
                };
                landingCells.Add((scanMinX, scanMaxX, fallbackSection));
                if (landingDirection >= 0)
                    positiveCapacity = (scanMaxX - scanMinX) * targetWidthFt;
                else
                    negativeCapacity = (scanMaxX - scanMinX) * targetWidthFt;
                debugLog?.Add("LandingConnector using run-end fallback because no footprint cell touched both route ends");
            }

            int resolvedLandingDirection = landingDirection;
            double capacityTolerance = tolFt * tolFt;
            double hintedCapacity = resolvedLandingDirection >= 0 ? positiveCapacity : negativeCapacity;
            double oppositeCapacity = resolvedLandingDirection >= 0 ? negativeCapacity : positiveCapacity;
            if (hintedCapacity <= capacityTolerance && oppositeCapacity > capacityTolerance)
                resolvedLandingDirection *= -1;

            List<(double leftX, double rightX, LandingSection2D section)> connectedCells = landingCells
                .Where(cell => (resolvedLandingDirection >= 0
                    ? cell.section.MaxY - innerAnchorY
                    : innerAnchorY - cell.section.MinY) > tolFt)
                .ToList();
            if (connectedCells.Count == 0)
                return false;

            double fullSpan = scanMaxX - scanMinX;
            double minStableSpan = Math.Min(MmToInternal(50.0), fullSpan * 0.1);
            var stableCells = connectedCells
                .Where(cell => cell.rightX - cell.leftX >= minStableSpan)
                .ToList();
            var widthCells = stableCells.Count > 0 ? stableCells : connectedCells;
            double narrowConnectedWidthFt = widthCells
                .Select(cell => resolvedLandingDirection >= 0
                    ? cell.section.MaxY - innerAnchorY
                    : innerAnchorY - cell.section.MinY)
                .Where(width => width > tolFt)
                .DefaultIfEmpty(targetWidthFt)
                .Min();
            double commonSectionWidthFt = Math.Min(targetWidthFt, narrowConnectedWidthFt);
            if (commonSectionWidthFt <= tolFt)
                return false;

            double commonMinY = resolvedLandingDirection >= 0
                ? innerAnchorY
                : innerAnchorY - commonSectionWidthFt;
            double commonMaxY = resolvedLandingDirection >= 0
                ? innerAnchorY + commonSectionWidthFt
                : innerAnchorY;

            int candidateCellCount = 0;
            int failedCellCount = 0;

            var connectorPolygon = new List<XYZ>
            {
                new XYZ(scanMinX, commonMinY, 0.0),
                new XYZ(scanMaxX, commonMinY, 0.0),
                new XYZ(scanMaxX, commonMaxY, 0.0),
                new XYZ(scanMinX, commonMaxY, 0.0)
            };

            candidateCellCount = 1;
            Solid connectorSolid = TryCreateVerticalExtrusionFromLocalPolygon2D(
                connectorPolygon,
                xDir,
                yDir,
                baseZ,
                heightFt,
                tolFt);
            if (connectorSolid != null && connectorSolid.Volume > 1e-9)
                routeSolids.Add(connectorSolid);
            else
                failedCellCount++;

            routeSolids = UnionConnectedLandingSolids(routeSolids, debugLog);

            debugLog?.Add(
                $"LandingSectionLimit={(routeSolids.Count > 0 ? "ok" : "failed")}; " +
                $"directionHint={landingDirection}; direction={resolvedLandingDirection}; targetWidth={FormatFtMm(targetWidthFt)}; " +
                $"commonWidth={FormatFtMm(commonSectionWidthFt)}; " +
                $"innerAnchorY={FormatFt(innerAnchorY)}; " +
                $"scanX={FormatFt(scanMinX)}..{FormatFt(scanMaxX)}; " +
                $"positiveCapacity={FormatFt(positiveCapacity)}; negativeCapacity={FormatFt(negativeCapacity)}; " +
                $"actualNarrowWidth={FormatFtMm(narrowConnectedWidthFt)}; " +
                $"cells={candidateCellCount}; solids={routeSolids.Count}; failed={failedCellCount}");

            return routeSolids.Count > 0;
        }

        private static List<Solid> UnionConnectedLandingSolids(List<Solid> sourceSolids, RouteDebugLog debugLog)
        {
            List<Solid> validSolids = GetValidSolids(sourceSolids);
            if (validSolids.Count <= 1)
                return validSolids;

            var mergedSolids = new List<Solid>();
            int unionCount = 0;

            foreach (Solid source in validSolids)
            {
                Solid current = source;
                bool merged;

                do
                {
                    merged = false;
                    for (int i = 0; i < mergedSolids.Count; i++)
                    {
                        try
                        {
                            Solid union = BooleanOperationsUtils.ExecuteBooleanOperation(
                                mergedSolids[i],
                                current,
                                BooleanOperationsType.Union);
                            if (union == null || union.Volume <= 1e-9)
                                continue;

                            current = union;
                            mergedSolids.RemoveAt(i);
                            unionCount++;
                            merged = true;
                            break;
                        }
                        catch
                        {
                        }
                    }
                }
                while (merged);

                mergedSolids.Add(current);
            }

            mergedSolids = GetValidSolids(mergedSolids);
            debugLog?.Add($"LandingSolidUnion input={validSolids.Count}; output={mergedSolids.Count}; unions={unionCount}");
            return mergedSolids;
        }
      
        private static bool TryGetLandingSectionClosestToY(
            List<List<XYZ>> footprintLoops,
            double x,
            double anchorY,
            double tolFt,
            out LandingSection2D section)
        {
            section = new LandingSection2D();

            LandingSection2D positiveSection;
            LandingSection2D negativeSection;
            bool hasPositive = TryGetOutermostLandingSectionAtX(
                footprintLoops,
                x,
                1,
                tolFt,
                out positiveSection);
            bool hasNegative = TryGetOutermostLandingSectionAtX(
                footprintLoops,
                x,
                -1,
                tolFt,
                out negativeSection);

            if (!hasPositive && !hasNegative)
                return false;
            if (!hasNegative)
            {
                section = positiveSection;
                return true;
            }
            if (!hasPositive)
            {
                section = negativeSection;
                return true;
            }

            double positiveDistance = GetDistanceToRange(anchorY, positiveSection.MinY, positiveSection.MaxY);
            double negativeDistance = GetDistanceToRange(anchorY, negativeSection.MinY, negativeSection.MaxY);
            section = positiveDistance <= negativeDistance ? positiveSection : negativeSection;
            return section.Width > tolFt;
        }

        private static bool TryGetOutermostLandingSectionAtX(
            List<List<XYZ>> footprintLoops,
            double x,
            int landingDirection,
            double tolFt,
            out LandingSection2D section)
        {
            section = new LandingSection2D();
            var ranges = new List<ProjectionRange2D>();

            foreach (List<XYZ> loop in footprintLoops ?? new List<List<XYZ>>())
            {
                if (loop == null || loop.Count < 3)
                    continue;

                var intersections = new List<double>();

                for (int i = 0; i < loop.Count; i++)
                {
                    XYZ a = loop[i];
                    XYZ b = loop[(i + 1) % loop.Count];
                    if (a == null || b == null || Math.Abs(b.X - a.X) <= tolFt)
                        continue;

                    double edgeMinX = Math.Min(a.X, b.X);
                    double edgeMaxX = Math.Max(a.X, b.X);
                    if (x < edgeMinX || x >= edgeMaxX)
                        continue;

                    double t = (x - a.X) / (b.X - a.X);
                    intersections.Add(a.Y + (b.Y - a.Y) * t);
                }

                intersections = intersections
                    .OrderBy(y => y)
                    .Aggregate(new List<double>(), (result, value) =>
                    {
                        if (result.Count == 0 || Math.Abs(value - result[result.Count - 1]) > tolFt)
                            result.Add(value);
                        return result;
                    });

                for (int i = 0; i + 1 < intersections.Count; i += 2)
                {
                    double minY = intersections[i];
                    double maxY = intersections[i + 1];
                    if (maxY - minY <= tolFt)
                        continue;

                    ranges.Add(new ProjectionRange2D { MinY = minY, MaxY = maxY });
                }
            }

            if (ranges.Count == 0)
                return false;

            ProjectionRange2D selected = landingDirection >= 0
                ? ranges.OrderByDescending(r => r.MaxY).First()
                : ranges.OrderBy(r => r.MinY).First();

            section = new LandingSection2D
            {
                X = x,
                MinY = selected.MinY,
                MaxY = selected.MaxY
            };
            return section.Width > tolFt;
        }

        private static bool TrySelectLandingCandidateAxes(
            RunRouteBodyInfo runA,
            RunRouteBodyInfo runB,
            XYZ faceCenterA,
            XYZ faceCenterB,
            XYZ fallbackXDir,
            XYZ fallbackYDir,
            double thresholdFt,
            out XYZ xDir,
            out XYZ yDir,
            out double dx,
            out double dy,
            out string source)
        {
            xDir = fallbackXDir;
            yDir = fallbackYDir;
            dx = 0.0;
            dy = 0.0;
            source = "fallback";

            var candidates = new List<Tuple<string, XYZ, XYZ>>
            {
                Tuple.Create("A.run", runA?.XDirPlan, runA?.YDirPlan),
                Tuple.Create("B.run", runB?.XDirPlan, runB?.YDirPlan),
                Tuple.Create("fallback", fallbackXDir, fallbackYDir)
            };

            double bestScore = double.NegativeInfinity;
            XYZ delta = new XYZ(faceCenterB.X - faceCenterA.X, faceCenterB.Y - faceCenterA.Y, 0);

            foreach (var candidate in candidates)
            {
                XYZ cx;
                XYZ cy;
                if (!TryNormalizePlanAxes(candidate.Item2, candidate.Item3, out cx, out cy))
                    continue;

                double candidateDx = Math.Abs(delta.DotProduct(cx));
                double candidateDy = Math.Abs(delta.DotProduct(cy));
                double score = Math.Min(candidateDx, candidateDy);
                if (score <= bestScore)
                    continue;

                bestScore = score;
                xDir = cx;
                yDir = cy;
                dx = candidateDx;
                dy = candidateDy;
                source = candidate.Item1;
            }

            return dx > thresholdFt && dy > thresholdFt;
        }

        private static bool TryNormalizePlanAxes(XYZ rawX, XYZ rawY, out XYZ xDir, out XYZ yDir)
        {
            xDir = null;
            yDir = null;

            if (rawX == null || rawX.GetLength() < 1e-9)
                return false;

            xDir = new XYZ(rawX.X, rawX.Y, 0);
            if (xDir.GetLength() < 1e-9)
                return false;

            xDir = xDir.Normalize();

            if (rawY != null && rawY.GetLength() > 1e-9)
                yDir = new XYZ(rawY.X, rawY.Y, 0);

            if (yDir == null || yDir.GetLength() < 1e-9)
                yDir = XYZ.BasisZ.CrossProduct(xDir);

            yDir = yDir - xDir * yDir.DotProduct(xDir);
            if (yDir.GetLength() < 1e-9)
                yDir = XYZ.BasisZ.CrossProduct(xDir);

            if (yDir.GetLength() < 1e-9)
                return false;

            yDir = yDir.Normalize();
            return true;
        }

        private static XYZ ProjectLocal2D(XYZ point, XYZ xDir, XYZ yDir)
        {
            XYZ flat = new XYZ(point.X, point.Y, 0);
            return new XYZ(flat.DotProduct(xDir), flat.DotProduct(yDir), 0);
        }

        private static void EnsureSpanAtLeast(ref double min, ref double max, double center, double minSize)
        {
            if (max - min >= minSize)
                return;

            double half = minSize * 0.5;
            min = center - half;
            max = center + half;
        }

        private static LocalRect2D ExpandLocalRect(LocalRect2D rect, double amount)
        {
            return new LocalRect2D(rect.MinX - amount, rect.MaxX + amount, rect.MinY - amount, rect.MaxY + amount, rect.Name, rect.Group);
        }

        private static bool TryCreateLandingRouteSolidsFromFootprint(
            StairsLanding landing,
            XYZ xDir,
            XYZ yDir,
            List<LocalRect2D> candidateRects,
            double baseZ,
            double heightFt,
            RouteDebugLog debugLog,
            out List<Solid> routeSolids)
        {
            routeSolids = new List<Solid>();
            if (landing == null || xDir == null || yDir == null || candidateRects == null || candidateRects.Count == 0 || heightFt <= 1e-9)
                return false;

            List<List<XYZ>> footprintLoops;
            if (!TryGetLandingTopFootprintLoops2D(landing, xDir, yDir, debugLog, out footprintLoops))
                return false;

            double tolFt = MmToInternal(1.0);
            string bestGroup = "";
            int bestClipped = 0;
            int bestFailed = 0;
            double bestVolume = double.NegativeInfinity;
            var bestSolids = new List<Solid>();

            foreach (var group in candidateRects.GroupBy(x => string.IsNullOrWhiteSpace(x.Group) ? x.Name : x.Group))
            {
                var groupSolids = new List<Solid>();
                int clipped = 0;
                int failed = 0;

                foreach (LocalRect2D rect in group)
                {
                    foreach (List<XYZ> footprintLoop in footprintLoops)
                    {
                        List<XYZ> clippedPolygon = ClipPolygonToRect2D(footprintLoop, rect.MinX, rect.MaxX, rect.MinY, rect.MaxY, tolFt);
                        clippedPolygon = CleanPolygon2D(clippedPolygon, tolFt);
                        if (clippedPolygon.Count < 3 || Math.Abs(GetSignedPolygonArea2D(clippedPolygon)) <= tolFt * tolFt)
                            continue;

                        clipped++;
                        Solid solid = TryCreateVerticalExtrusionFromLocalPolygon2D(clippedPolygon, xDir, yDir, baseZ, heightFt, tolFt);
                        if (solid != null && solid.Volume > 1e-9)
                            groupSolids.Add(solid);
                        else
                            failed++;
                    }
                }

                groupSolids = GetValidSolids(groupSolids);
                double volume = SumSolidVolumes(groupSolids);
                debugLog?.Add($"LandingFootprintClip candidate={group.Key}; rects={group.Count()}; clipped={clipped}; solids={groupSolids.Count}; failed={failed}; volume={FormatFt(volume)}");

                if (volume <= bestVolume)
                    continue;

                bestVolume = volume;
                bestGroup = group.Key;
                bestClipped = clipped;
                bestFailed = failed;
                bestSolids = groupSolids;
            }

            routeSolids = GetValidSolids(bestSolids);
            debugLog?.Add($"LandingFootprintClip=ok; selected={bestGroup}; footprintLoops={footprintLoops.Count}; candidateGroups={candidateRects.Select(x => x.Group).Distinct().Count()}; clipped={bestClipped}; solids={routeSolids.Count}; failed={bestFailed}; volume={FormatFt(bestVolume)}");

            return routeSolids.Count > 0;
        }

        private static bool TryGetLandingTopFootprintLoops2D(StairsLanding landing, XYZ xDir, XYZ yDir, RouteDebugLog debugLog, out List<List<XYZ>> loops2D)
        {
            loops2D = new List<List<XYZ>>();
            if (landing == null)
                return false;

            var solids = new List<Solid>();
            AddElementSolids(landing, solids);
            if (solids.Count == 0)
            {
                debugLog?.Add("LandingFootprint2D=no landing solids");
                return false;
            }

            double topZ = double.NegativeInfinity;
            foreach (Solid solid in solids)
            {
                if (solid == null || solid.Volume < 1e-9)
                    continue;

                foreach (Face face in solid.Faces)
                {
                    PlanarFace pf = face as PlanarFace;
                    if (pf == null)
                        continue;

                    XYZ normal = pf.FaceNormal;
                    if (normal == null || normal.Z < 0.95)
                        continue;

                    if (pf.Origin.Z > topZ)
                        topZ = pf.Origin.Z;
                }
            }

            if (double.IsNegativeInfinity(topZ))
            {
                debugLog?.Add("LandingFootprint2D=no top faces");
                return false;
            }

            double topToleranceFt = MmToInternal(8.0);
            double tolFt = MmToInternal(1.0);
            var seen = new HashSet<string>();
            int faceCount = 0;

            foreach (Solid solid in solids)
            {
                if (solid == null || solid.Volume < 1e-9)
                    continue;

                foreach (Face face in solid.Faces)
                {
                    PlanarFace pf = face as PlanarFace;
                    if (pf == null)
                        continue;

                    XYZ normal = pf.FaceNormal;
                    if (normal == null || normal.Z < 0.95)
                        continue;

                    if (pf.Origin.Z < topZ - topToleranceFt)
                        continue;

                    faceCount++;

                    IList<CurveLoop> curveLoops;
                    try
                    {
                        curveLoops = pf.GetEdgesAsCurveLoops();
                    }
                    catch (Exception ex)
                    {
                        debugLog?.Add($"LandingFootprint2D=skip face; loop error: {ex.Message}");
                        continue;
                    }

                    var faceLoops2D = new List<List<XYZ>>();
                    foreach (CurveLoop curveLoop in curveLoops ?? new List<CurveLoop>())
                    {
                        List<XYZ> loop2D = CurveLoopToLocalPolygon2D(curveLoop, xDir, yDir, tolFt);
                        loop2D = CleanPolygon2D(loop2D, tolFt);
                        if (loop2D.Count < 3)
                            continue;

                        double area = Math.Abs(GetSignedPolygonArea2D(loop2D));
                        if (area <= tolFt * tolFt)
                            continue;

                        faceLoops2D.Add(loop2D);
                    }

                    foreach (List<XYZ> loop2D in faceLoops2D.OrderByDescending(x => Math.Abs(GetSignedPolygonArea2D(x))).Take(1))
                    {
                        string signature = GetPolygonSignature2D(loop2D);
                        if (!seen.Add(signature))
                            continue;

                        loops2D.Add(loop2D);
                    }
                }
            }

            loops2D = loops2D
                .OrderByDescending(x => Math.Abs(GetSignedPolygonArea2D(x)))
                .ToList();

            debugLog?.Add($"LandingFootprint2D topZ={FormatFt(topZ)}; faces={faceCount}; loops={loops2D.Count}");
            return loops2D.Count > 0;
        }

        private static List<XYZ> CurveLoopToLocalPolygon2D(CurveLoop curveLoop, XYZ xDir, XYZ yDir, double tolFt)
        {
            var result = new List<XYZ>();
            if (curveLoop == null)
                return result;

            foreach (Curve curve in curveLoop)
            {
                if (curve == null)
                    continue;

                IList<XYZ> tessellated = null;
                try { tessellated = curve.Tessellate(); } catch { }

                if (tessellated == null || tessellated.Count == 0)
                {
                    try
                    {
                        tessellated = new List<XYZ> { curve.GetEndPoint(0), curve.GetEndPoint(1) };
                    }
                    catch
                    {
                        continue;
                    }
                }

                foreach (XYZ p in tessellated)
                {
                    if (p == null)
                        continue;

                    XYZ flat = new XYZ(p.X, p.Y, 0);
                    XYZ local = new XYZ(flat.DotProduct(xDir), flat.DotProduct(yDir), 0);
                    if (result.Count == 0 || Distance2D(result[result.Count - 1], local) > tolFt)
                        result.Add(local);
                }
            }

            return CleanPolygon2D(result, tolFt);
        }

        private static List<XYZ> ClipPolygonToRect2D(List<XYZ> polygon, double minX, double maxX, double minY, double maxY, double tolFt)
        {
            List<XYZ> result = CleanPolygon2D(polygon, tolFt);
            result = ClipPolygonByBoundary2D(result, p => p.X >= minX - tolFt, (a, b) => IntersectSegmentAtX2D(a, b, minX));
            result = ClipPolygonByBoundary2D(result, p => p.X <= maxX + tolFt, (a, b) => IntersectSegmentAtX2D(a, b, maxX));
            result = ClipPolygonByBoundary2D(result, p => p.Y >= minY - tolFt, (a, b) => IntersectSegmentAtY2D(a, b, minY));
            result = ClipPolygonByBoundary2D(result, p => p.Y <= maxY + tolFt, (a, b) => IntersectSegmentAtY2D(a, b, maxY));
            return CleanPolygon2D(result, tolFt);
        }

        private static List<XYZ> ClipPolygonByBoundary2D(List<XYZ> input, Func<XYZ, bool> inside, Func<XYZ, XYZ, XYZ> intersection)
        {
            var output = new List<XYZ>();
            if (input == null || input.Count == 0)
                return output;

            XYZ previous = input[input.Count - 1];
            bool previousInside = inside(previous);

            foreach (XYZ current in input)
            {
                bool currentInside = inside(current);

                if (currentInside)
                {
                    if (!previousInside)
                        output.Add(intersection(previous, current));
                    output.Add(current);
                }
                else if (previousInside)
                {
                    output.Add(intersection(previous, current));
                }

                previous = current;
                previousInside = currentInside;
            }

            return output;
        }

        private static XYZ IntersectSegmentAtX2D(XYZ a, XYZ b, double x)
        {
            double dx = b.X - a.X;
            if (Math.Abs(dx) < 1e-12)
                return new XYZ(x, (a.Y + b.Y) * 0.5, 0);

            double t = (x - a.X) / dx;
            t = Math.Max(0.0, Math.Min(1.0, t));
            return new XYZ(x, a.Y + (b.Y - a.Y) * t, 0);
        }

        private static XYZ IntersectSegmentAtY2D(XYZ a, XYZ b, double y)
        {
            double dy = b.Y - a.Y;
            if (Math.Abs(dy) < 1e-12)
                return new XYZ((a.X + b.X) * 0.5, y, 0);

            double t = (y - a.Y) / dy;
            t = Math.Max(0.0, Math.Min(1.0, t));
            return new XYZ(a.X + (b.X - a.X) * t, y, 0);
        }

        private static Solid TryCreateVerticalExtrusionFromLocalPolygon2D(List<XYZ> polygon, XYZ xDir, XYZ yDir, double baseZ, double heightFt, double tolFt)
        {
            List<XYZ> cleaned = CleanPolygon2D(polygon, tolFt);
            if (cleaned.Count < 3)
                return null;

            if (GetSignedPolygonArea2D(cleaned) < 0.0)
                cleaned.Reverse();

            Solid solid = TryCreateVerticalExtrusionFromOrderedLocalPolygon2D(cleaned, xDir, yDir, baseZ, heightFt, tolFt);
            if (solid != null && solid.Volume > 1e-9)
                return solid;

            cleaned.Reverse();
            return TryCreateVerticalExtrusionFromOrderedLocalPolygon2D(cleaned, xDir, yDir, baseZ, heightFt, tolFt);
        }

        private static Solid TryCreateVerticalExtrusionFromOrderedLocalPolygon2D(List<XYZ> polygon, XYZ xDir, XYZ yDir, double baseZ, double heightFt, double tolFt)
        {
            if (polygon == null || polygon.Count < 3)
                return null;

            try
            {
                var loop = new CurveLoop();
                for (int i = 0; i < polygon.Count; i++)
                {
                    XYZ a = Local2DToWorld(polygon[i], xDir, yDir, baseZ);
                    XYZ b = Local2DToWorld(polygon[(i + 1) % polygon.Count], xDir, yDir, baseZ);
                    if (a.DistanceTo(b) <= tolFt)
                        continue;

                    loop.Append(Line.CreateBound(a, b));
                }

                return GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { loop }, XYZ.BasisZ, heightFt);
            }
            catch
            {
                return null;
            }
        }

        private static XYZ Local2DToWorld(XYZ local, XYZ xDir, XYZ yDir, double z)
        {
            return xDir * local.X + yDir * local.Y + XYZ.BasisZ * z;
        }

        private static List<XYZ> CleanPolygon2D(IEnumerable<XYZ> points, double tolFt)
        {
            var result = new List<XYZ>();
            foreach (XYZ point in points ?? Enumerable.Empty<XYZ>())
            {
                if (point == null)
                    continue;

                XYZ p = new XYZ(point.X, point.Y, 0);
                if (result.Count == 0 || Distance2D(result[result.Count - 1], p) > tolFt)
                    result.Add(p);
            }

            if (result.Count > 1 && Distance2D(result[0], result[result.Count - 1]) <= tolFt)
                result.RemoveAt(result.Count - 1);

            bool changed = true;
            while (changed && result.Count >= 3)
            {
                changed = false;
                for (int i = 0; i < result.Count; i++)
                {
                    XYZ prev = result[(i + result.Count - 1) % result.Count];
                    XYZ current = result[i];
                    XYZ next = result[(i + 1) % result.Count];
                    double len = Distance2D(prev, current) + Distance2D(current, next);
                    double cross = Math.Abs(Cross2D(current - prev, next - current));

                    if (Distance2D(prev, current) <= tolFt || Distance2D(current, next) <= tolFt || cross <= tolFt * Math.Max(len, tolFt))
                    {
                        result.RemoveAt(i);
                        changed = true;
                        break;
                    }
                }
            }

            return result;
        }

        private static double GetSignedPolygonArea2D(IList<XYZ> points)
        {
            if (points == null || points.Count < 3)
                return 0.0;

            double area = 0.0;
            for (int i = 0; i < points.Count; i++)
            {
                XYZ a = points[i];
                XYZ b = points[(i + 1) % points.Count];
                area += a.X * b.Y - b.X * a.Y;
            }

            return area * 0.5;
        }

        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static string GetPolygonSignature2D(IList<XYZ> points)
        {
            if (points == null || points.Count == 0)
                return string.Empty;

            return string.Join(";", points.Select(p => $"{Math.Round(p.X, 4)}:{Math.Round(p.Y, 4)}"));
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

        private static void AddElementAndDependentSolids(Element elem, List<Solid> solids)
        {
            AddElementAndDependentSolidsRecursive(elem, solids, new HashSet<long>(), 0);
        }

        private static void AddElementAndDependentSolidsRecursive(
            Element elem,
            List<Solid> solids,
            HashSet<long> visited,
            int depth)
        {
            if (elem == null || solids == null || visited == null || depth > 3)
                return;

            long id = IDHelper.ElIdValue(elem.Id);
            if (id > 0 && !visited.Add(id))
                return;

            AddElementSolids(elem, solids);

            ICollection<ElementId> dependentIds = null;
            try { dependentIds = elem.GetDependentElements(null); }
            catch { }

            if (dependentIds == null || dependentIds.Count == 0)
                return;

            Document doc = elem.Document;
            foreach (ElementId dependentId in dependentIds)
            {
                Element dependent = null;
                try { dependent = doc?.GetElement(dependentId); }
                catch { }

                AddElementAndDependentSolidsRecursive(
                    dependent,
                    solids,
                    visited,
                    depth + 1);
            }
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

        private static bool IsEvacuationRouteShape(Element elem)
        {
            DirectShape ds = elem as DirectShape;
            if (ds == null || !IsOwnRouteShape(ds))
                return false;

            string name = "";
            try { name = ds.Name ?? ""; } catch { }
            return name.StartsWith("ПЭ", StringComparison.OrdinalIgnoreCase);
        }

        private static long GetRouteComponentElementId(DirectShape routeShape)
        {
            if (routeShape == null)
                return 0;

            string appDataId = "";
            try { appDataId = routeShape.ApplicationDataId ?? ""; } catch { }

            long componentId = TryParseLastElementIdToken(appDataId);
            if (componentId > 0)
                return componentId;

            string name = "";
            try { name = routeShape.Name ?? ""; } catch { }
            return TryParseLastElementIdToken(name);
        }

        private static long TryParseLastElementIdToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            MatchCollection matches = Regex.Matches(text, @"\d+");
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                if (long.TryParse(matches[i].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) && value > 0)
                    return value;
            }

            return 0;
        }

        private static long FindOwnerStairIdByComponentId(Document doc, long componentId)
        {
            if (doc == null || componentId <= 0)
                return 0;

            foreach (Stairs stairs in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .OfType<Stairs>())
            {
                try
                {
                    if ((stairs.GetStairsRuns() ?? new List<ElementId>()).Any(x => IDHelper.ElIdValue(x) == componentId))
                        return IDHelper.ElIdValue(stairs.Id);

                    if ((stairs.GetStairsLandings() ?? new List<ElementId>()).Any(x => IDHelper.ElIdValue(x) == componentId))
                        return IDHelper.ElIdValue(stairs.Id);
                }
                catch
                {
                }
            }

            return 0;
        }

        private static double MmToInternal(double valueMm)
        {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double RoundWidthDownTo5Mm(double widthFt)
        {
            if (widthFt <= 1e-9)
                return widthFt;

            double widthMm = IDHelper.ConvertInternalToMm(widthFt);
            double roundedMm = Math.Floor((widthMm + 1e-6) / 5.0) * 5.0;
            return roundedMm > 0.0 ? MmToInternal(roundedMm) : widthFt;
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

        private static RunClearWidthInfo GetRunClearWidthInfo(Document doc, Stairs stairs, StairsRun run, XYZ bottom, XYZ top, XYZ xP, XYZ yP, double lenPlan, bool considerRailings, RouteDebugLog debugLog)
        {
            double pathCenterY = new XYZ(bottom.X, bottom.Y, 0).DotProduct(yP);
            double nominalMinY;
            double nominalMaxY;
            double nominalWidthFt;

            if (TryGetRunGeometryWidthRange(run, xP, yP, bottom, top, lenPlan, out nominalMinY, out nominalMaxY))
            {
                double sideAllowanceFt = GetRunWidthSideAllowanceFt();
                nominalMinY -= sideAllowanceFt;
                nominalMaxY += sideAllowanceFt;
                nominalWidthFt = nominalMaxY - nominalMinY;
                debugLog?.Add($"RunWidthSource=Geometry; sideAllowance={FormatFtMm(sideAllowanceFt)} minY={FormatFt(nominalMinY)} maxY={FormatFt(nominalMaxY)} width={FormatFtMm(nominalWidthFt)} pathCenterY={FormatFt(pathCenterY)}");
            }
            else
            {
                nominalWidthFt = GetRunWidthFt(run, bottom, top, xP, yP);
                nominalMinY = pathCenterY - nominalWidthFt / 2.0;
                nominalMaxY = pathCenterY + nominalWidthFt / 2.0;
                debugLog?.Add($"RunWidthSource=Fallback; minY={FormatFt(nominalMinY)} maxY={FormatFt(nominalMaxY)} width={FormatFtMm(nominalWidthFt)} pathCenterY={FormatFt(pathCenterY)}");
            }

            if (TryGetActualRunWidthFt(run, out double actualRunWidthFt) && actualRunWidthFt > 1e-9)
            {
                double nominalCenterY = (nominalMinY + nominalMaxY) * 0.5;
                double widthDeltaFt = Math.Abs(nominalWidthFt - actualRunWidthFt);

                nominalWidthFt = actualRunWidthFt;
                nominalMinY = nominalCenterY - nominalWidthFt / 2.0;
                nominalMaxY = nominalCenterY + nominalWidthFt / 2.0;

                if (widthDeltaFt > MmToInternal(1.0))
                    debugLog?.Add($"RunWidthSource=ActualRunWidth override; actual={FormatFtMm(actualRunWidthFt)} delta={FormatFtMm(widthDeltaFt)} minY={FormatFt(nominalMinY)} maxY={FormatFt(nominalMaxY)}");
                else
                    debugLog?.Add($"RunWidthSource=ActualRunWidth confirmed; actual={FormatFtMm(actualRunWidthFt)}");
            }

            var result = new RunClearWidthInfo
            {
                WidthFt = nominalWidthFt,
                CenterOffsetFt = (nominalMinY + nominalMaxY) * 0.5 - pathCenterY,
                ClearMinY = nominalMinY,
                ClearMaxY = nominalMaxY,
                HasRailingBoundary = false,
                HasLeftRailingBoundary = false,
                HasRightRailingBoundary = false
            };

            if (doc == null || run == null || nominalWidthFt <= 1e-9 || lenPlan <= 1e-9)
                return result;

            BoundingBoxXYZ runBox = run.get_BoundingBox(null);
            if (runBox == null)
                return result;

            double clearMinY = nominalMinY;
            double clearMaxY = nominalMaxY;

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

            Solid obstacleTestSolid = TryCreateRunClearWidthObstacleTestSolid(run, bottom, top, yP, nominalWidthFt, debugLog);
            double railingSideSearchFt = MmToInternal(300.0);
            Solid railingObstacleTestSolid = TryCreateRunClearWidthObstacleTestSolid(
                run,
                bottom,
                top,
                yP,
                nominalWidthFt + railingSideSearchFt * 2.0,
                debugLog);

            var excludedIds = GetStairAndComponentIds(stairs, includeAssociatedRailings: false);
            excludedIds.Add(IDHelper.ElIdValue(run.Id));
            debugLog?.Add($"RunClearWidth excluded current stair/components count={excludedIds.Count}");

            var candidatesById = new Dictionary<long, Element>();
            try
            {
                foreach (Element candidate in new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementMulticategoryFilter(categories))
                    .WherePasses(new BoundingBoxIntersectsFilter(outline))
                    .ToElements())
                {
                    if (candidate != null)
                        candidatesById[IDHelper.ElIdValue(candidate.Id)] = candidate;
                }
            }
            catch
            {
            }

            try
            {
                foreach (Element railing in new FilteredElementCollector(doc)
                    .OfClass(typeof(Railing))
                    .WherePasses(new BoundingBoxIntersectsFilter(outline))
                    .ToElements())
                {
                    if (railing != null)
                        candidatesById[IDHelper.ElIdValue(railing.Id)] = railing;
                }
            }
            catch
            {
            }

            IEnumerable<Element> candidates = candidatesById.Values;

            double minUsableWidthFt = MmToInternal(300.0);
            double referenceCenterY = (nominalMinY + nominalMaxY) * 0.5;
            double railingClearMinY = nominalMinY - railingSideSearchFt;
            double railingClearMaxY = nominalMaxY + railingSideSearchFt;
            bool hasLeftRailingBoundary = false;
            bool hasRightRailingBoundary = false;

            foreach (Element elem in candidates)
            {
                if (elem == null) continue;
                if (excludedIds.Contains(IDHelper.ElIdValue(elem.Id))) continue;
                if (IsOwnRouteShape(elem)) continue;

                bool isRailing = elem is Railing || IsRailingObstacle(elem);
                if (isRailing && !considerRailings)
                    continue;

                if (isRailing && TryGetRailingPathProjectionRanges(elem, xP, yP, out List<ProjectionRange2D> railingPathRanges))
                {
                    Solid railingOverlapSolid = railingObstacleTestSolid ?? obstacleTestSolid;
                    bool hasSolidOverlap = railingOverlapSolid != null
                        && HasClearWidthObstacleSolidOverlap(railingOverlapSolid, elem);
                    bool hasPathProximity = HasRailingPathNearRun(
                        railingPathRanges,
                        runX0,
                        runX1,
                        nominalMinY,
                        nominalMaxY,
                        railingSideSearchFt);
                    if (!hasSolidOverlap && !hasPathProximity)
                    {
                        debugLog?.Add($"RunClearWidth railing path skipped no-run-overlap ID={IDHelper.ElIdValue(elem.Id)} host={GetRailingHostDebugValue(elem)} ranges={railingPathRanges.Count}");
                        continue;
                    }

                    bool hasPhysicalLeft;
                    bool hasPhysicalRight;
                    ApplyRailingPhysicalSideBoundaries(
                        elem,
                        xP,
                        yP,
                        runX0,
                        runX1,
                        referenceCenterY,
                        nominalMinY,
                        nominalMaxY,
                        railingSideSearchFt,
                        debugLog,
                        ref railingClearMinY,
                        ref railingClearMaxY,
                        out hasPhysicalLeft,
                        out hasPhysicalRight);

                    if (hasPhysicalLeft)
                        hasLeftRailingBoundary = true;
                    if (hasPhysicalRight)
                        hasRightRailingBoundary = true;

                    foreach (ProjectionRange2D range in railingPathRanges)
                    {
                        bool pathOnLeft = range.MaxY <= referenceCenterY
                            && range.MinY < referenceCenterY;
                        bool pathOnRight = range.MinY >= referenceCenterY
                            && range.MaxY > referenceCenterY;

                        if ((pathOnLeft && hasPhysicalLeft)
                            || (pathOnRight && hasPhysicalRight))
                        {
                            continue;
                        }

                        bool pathApplied = ApplyRunClearWidthObstacleRange(
                            elem,
                            range.MinX,
                            range.MaxX,
                            range.MinY,
                            range.MaxY,
                            runX0,
                            runX1,
                            referenceCenterY,
                            minUsableWidthFt,
                            debugLog,
                            ref railingClearMinY,
                            ref railingClearMaxY,
                            "path");
                        if (pathApplied)
                        {
                            if (pathOnLeft)
                                hasLeftRailingBoundary = true;
                            if (pathOnRight)
                                hasRightRailingBoundary = true;
                        }
                    }

                    debugLog?.Add($"RunClearWidth railing paths processed ID={IDHelper.ElIdValue(elem.Id)} host={GetRailingHostDebugValue(elem)} ranges={railingPathRanges.Count} solidOverlap={hasSolidOverlap} pathProximity={hasPathProximity} physicalLeft={hasPhysicalLeft} physicalRight={hasPhysicalRight}");
                    continue;
                }

                var solids = new List<Solid>();
                AddElementSolids(elem, solids);
                foreach (Solid solid in GetValidSolids(solids))
                {
                    if (!TryGetClearWidthObstacleProjectionRange(obstacleTestSolid, elem, solid, xP, yP, out double minX, out double maxX, out double minY, out double maxY))
                    {
                        debugLog?.Add($"RunClearWidth obstacle skipped no-solid-overlap ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}'");
                        continue;
                    }

                    if (isRailing)
                    {
                        bool railingApplied = ApplyRunClearWidthObstacleRange(
                            elem,
                            minX,
                            maxX,
                            minY,
                            maxY,
                            runX0,
                            runX1,
                            referenceCenterY,
                            minUsableWidthFt,
                            debugLog,
                            ref railingClearMinY,
                            ref railingClearMaxY,
                            "solid-fallback");
                        if (railingApplied)
                        {
                            if (maxY <= referenceCenterY && minY < referenceCenterY)
                                hasLeftRailingBoundary = true;
                            if (minY >= referenceCenterY && maxY > referenceCenterY)
                                hasRightRailingBoundary = true;
                        }
                    }
                    else
                    {
                        ApplyRunClearWidthObstacleRange(
                            elem,
                            minX,
                            maxX,
                            minY,
                            maxY,
                            runX0,
                            runX1,
                            referenceCenterY,
                            minUsableWidthFt,
                            debugLog,
                            ref clearMinY,
                            ref clearMaxY,
                            "solid");
                    }
                }
            }

            if (considerRailings && (hasLeftRailingBoundary || hasRightRailingBoundary))
            {
                double effectiveRailingMinY;
                double effectiveRailingMaxY;

                if (hasLeftRailingBoundary && hasRightRailingBoundary)
                {
                    effectiveRailingMinY = railingClearMinY;
                    effectiveRailingMaxY = railingClearMaxY;
                }
                else if (hasLeftRailingBoundary)
                {
                    effectiveRailingMinY = Math.Max(clearMinY, railingClearMinY);
                    effectiveRailingMaxY = clearMaxY;
                }
                else
                {
                    effectiveRailingMinY = clearMinY;
                    effectiveRailingMaxY = Math.Min(clearMaxY, railingClearMaxY);
                }

                double railingWidthFt = effectiveRailingMaxY - effectiveRailingMinY;
                if (railingWidthFt >= minUsableWidthFt)
                {
                    result.WidthFt = railingWidthFt;
                    result.CenterOffsetFt = (effectiveRailingMinY + effectiveRailingMaxY) * 0.5 - pathCenterY;
                    result.ClearMinY = effectiveRailingMinY;
                    result.ClearMaxY = effectiveRailingMaxY;
                    result.HasRailingBoundary = true;
                    result.HasLeftRailingBoundary = hasLeftRailingBoundary;
                    result.HasRightRailingBoundary = hasRightRailingBoundary;
                    debugLog?.Add(
                        $"RunRailingCorridor {(hasLeftRailingBoundary && hasRightRailingBoundary ? "complete" : "one-sided")} " +
                        $"left={hasLeftRailingBoundary} right={hasRightRailingBoundary} " +
                        $"width={FormatFtMm(railingWidthFt)} " +
                        $"centerOffset={FormatFtMm(result.CenterOffsetFt)} " +
                        $"clearY={FormatFt(effectiveRailingMinY)}..{FormatFt(effectiveRailingMaxY)}");
                    return result;
                }

                debugLog?.Add(
                    $"RunRailingCorridor rejected left={hasLeftRailingBoundary} right={hasRightRailingBoundary} " +
                    $"width={FormatFtMm(railingWidthFt)} min={FormatFtMm(minUsableWidthFt)}");
            }
            else if (considerRailings)
            {
                debugLog?.Add("RunRailingCorridor not found; using width mode fallback");
            }

            double clearWidthFt = clearMaxY - clearMinY;
            if (clearWidthFt < minUsableWidthFt || clearWidthFt > nominalWidthFt)
                return result;

            result.ClearMinY = clearMinY;
            result.ClearMaxY = clearMaxY;
            result.HasRailingBoundary = false;
            result.HasLeftRailingBoundary = false;
            result.HasRightRailingBoundary = false;

            if (nominalWidthFt - clearWidthFt < MmToInternal(1.0))
                return result;

            result.WidthFt = clearWidthFt;
            result.CenterOffsetFt = (clearMinY + clearMaxY) * 0.5 - pathCenterY;
            debugLog?.Add($"RunClearWidth result width={FormatFtMm(result.WidthFt)} centerOffset={FormatFtMm(result.CenterOffsetFt)} clearY={FormatFt(clearMinY)}..{FormatFt(clearMaxY)} nominal={FormatFtMm(nominalWidthFt)}");
            return result;
        }

        private static double GetRunManualWidthCenterOffset(StairsRun run, XYZ bottom, XYZ top, XYZ xP, XYZ yP, double lenPlan, RouteDebugLog debugLog)
        {
            if (run == null || bottom == null || top == null || xP == null || yP == null)
                return 0.0;

            double pathCenterY = new XYZ(bottom.X, bottom.Y, 0).DotProduct(yP);
            if (TryGetRunGeometryWidthRange(run, xP, yP, bottom, top, lenPlan, out double minY, out double maxY))
            {
                double geometryCenterY = (minY + maxY) * 0.5;
                double offset = geometryCenterY - pathCenterY;
                debugLog?.Add($"ManualWidthCenter=Geometry; minY={FormatFt(minY)} maxY={FormatFt(maxY)} offset={FormatFtMm(offset)}");
                return offset;
            }

            debugLog?.Add("ManualWidthCenter=Path; geometry range not found");
            return 0.0;
        }

        private static double GetMinRunWidthObstacleOverlapFt(double runLengthFt)
        {
            if (runLengthFt <= 1e-9)
                return MmToInternal(200.0);

            double byRatio = runLengthFt * 0.20;
            double min = MmToInternal(200.0);
            double max = MmToInternal(600.0);
            return Math.Max(min, Math.Min(max, byRatio));
        }

        private static bool ApplyRunClearWidthObstacleRange(
            Element elem,
            double minX,
            double maxX,
            double minY,
            double maxY,
            double runX0,
            double runX1,
            double referenceCenterY,
            double minUsableWidthFt,
            RouteDebugLog debugLog,
            ref double clearMinY,
            ref double clearMaxY,
            string source)
        {
            double overlapX = Math.Min(maxX, runX1) - Math.Max(minX, runX0);
            if (overlapX <= 0.0)
                return false;

            double minObstacleOverlapFt = IsRailingObstacle(elem)
                ? GetMinRailingPathObstacleOverlapFt(runX1 - runX0)
                : GetMinRunWidthObstacleOverlapFt(runX1 - runX0);

            if (overlapX < minObstacleOverlapFt)
            {
                debugLog?.Add($"RunClearWidth obstacle skipped {source} ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}' overlapX={FormatFtMm(overlapX)} minOverlap={FormatFtMm(minObstacleOverlapFt)} y={FormatFt(minY)}..{FormatFt(maxY)}");
                return false;
            }

            double sideToleranceFt = GetRunClearWidthSideToleranceFt(elem);
            bool fromLeft = maxY <= referenceCenterY && maxY > clearMinY - sideToleranceFt && minY < referenceCenterY;
            bool fromRight = minY >= referenceCenterY && minY < clearMaxY + sideToleranceFt && maxY > referenceCenterY;
            bool crossesCenter = minY < referenceCenterY && maxY > referenceCenterY;
            bool applied = false;

            if (fromLeft)
            {
                clearMinY = Math.Max(clearMinY, Math.Min(maxY, clearMaxY));
                applied = true;
                debugLog?.Add($"RunClearWidth obstacle left {source} ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}' overlapX={FormatFtMm(overlapX)} y={FormatFt(minY)}..{FormatFt(maxY)}");
            }

            if (fromRight)
            {
                clearMaxY = Math.Min(clearMaxY, Math.Max(minY, clearMinY));
                applied = true;
                debugLog?.Add($"RunClearWidth obstacle right {source} ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}' overlapX={FormatFtMm(overlapX)} y={FormatFt(minY)}..{FormatFt(maxY)}");
            }

            if (crossesCenter)
            {
                double leftWidth = Math.Max(0.0, minY - clearMinY);
                double rightWidth = Math.Max(0.0, clearMaxY - maxY);

                if (leftWidth >= minUsableWidthFt || rightWidth >= minUsableWidthFt)
                {
                    if (rightWidth >= leftWidth)
                        clearMinY = Math.Max(clearMinY, maxY);
                    else
                        clearMaxY = Math.Min(clearMaxY, minY);

                    applied = true;
                    debugLog?.Add($"RunClearWidth obstacle center {source} ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}' overlapX={FormatFtMm(overlapX)} chose={(rightWidth >= leftWidth ? "right" : "left")} left={FormatFtMm(leftWidth)} right={FormatFtMm(rightWidth)}");
                }
                else
                {
                    debugLog?.Add($"RunClearWidth obstacle center {source} ID={IDHelper.ElIdValue(elem.Id)} cat='{GetElementCategoryName(elem)}' overlapX={FormatFtMm(overlapX)} ignored; left={FormatFtMm(leftWidth)} right={FormatFtMm(rightWidth)}");
                }
            }

            return applied;
        }

        private static double GetMinRailingPathObstacleOverlapFt(double runLengthFt)
        {
            if (runLengthFt <= 1e-9)
                return MmToInternal(100.0);

            double byRatio = runLengthFt * 0.10;
            double min = MmToInternal(100.0);
            double max = MmToInternal(300.0);
            return Math.Max(min, Math.Min(max, byRatio));
        }

        private static double GetRunClearWidthSideToleranceFt(Element elem)
        {
            return (elem is Railing || IsRailingObstacle(elem))
                ? MmToInternal(300.0)
                : MmToInternal(5.0);
        }

        private static bool IsRailingObstacle(Element elem)
        {
            if (elem is Railing)
                return true;

            return IsElementInBuiltInCategory(elem,
                "OST_Railings",
                "OST_StairsRailing",
                "OST_RailingSystem",
                "OST_RailingRail",
                "OST_RailingTopRail",
                "OST_RailingHandRail",
                "OST_RailingSupport");
        }

        private static bool IsElementInBuiltInCategory(Element elem, params string[] categoryNames)
        {
            Category cat = elem == null ? null : elem.Category;
            if (cat == null || categoryNames == null)
                return false;

            int categoryId = IDHelper.ElIdInt(cat.Id);
            foreach (string categoryName in categoryNames)
            {
                try
                {
                    var bic = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), categoryName);
                    if (categoryId == (int)bic)
                        return true;
                }
                catch
                {
                }
            }

            return false;
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

            return categories;
        }

        private static double GetClearWidthObstacleMinIntersectionThicknessFt(Element elem)
        {
            return IsRailingObstacle(elem) ? MmToInternal(5.0) : MmToInternal(2.0);
        }

        private static bool HasClearWidthObstacleSolidOverlap(Solid obstacleTestSolid, Element elem)
        {
            if (obstacleTestSolid == null || obstacleTestSolid.Volume < 1e-9 || elem == null)
                return false;

            var solids = new List<Solid>();
            AddElementSolids(elem, solids);
            if (solids.Count == 0)
                return false;

            double minVolumeFt3 = GetMinIntersectionVolumeFt3();
            double minThicknessFt = GetClearWidthObstacleMinIntersectionThicknessFt(elem);

            foreach (Solid solid in GetValidSolids(solids))
            {
                if (HasMeaningfulSolidIntersection(obstacleTestSolid, solid, minVolumeFt3, minThicknessFt))
                    return true;
            }

            return false;
        }

        private static string GetRailingHostDebugValue(Element elem)
        {
            Railing railing = elem as Railing;
            if (railing == null)
                return "нет";

            try
            {
                ElementId hostId = railing.HostId;
                long id = IDHelper.ElIdValue(hostId);
                return id > 0 ? id.ToString(CultureInfo.InvariantCulture) : "нет";
            }
            catch
            {
                return "ошибка";
            }
        }

        private static bool HasRailingPathNearRun(
            List<ProjectionRange2D> ranges,
            double runX0,
            double runX1,
            double nominalMinY,
            double nominalMaxY,
            double sideSearchFt)
        {
            if (ranges == null || ranges.Count == 0)
                return false;

            double minLongitudinalOverlapFt = GetMinRailingPathObstacleOverlapFt(runX1 - runX0);
            foreach (ProjectionRange2D range in ranges)
            {
                double overlapX = Math.Min(range.MaxX, runX1) - Math.Max(range.MinX, runX0);
                if (overlapX < minLongitudinalOverlapFt)
                    continue;

                double sideDistance = 0.0;
                if (range.MaxY < nominalMinY)
                    sideDistance = nominalMinY - range.MaxY;
                else if (range.MinY > nominalMaxY)
                    sideDistance = range.MinY - nominalMaxY;

                if (sideDistance <= sideSearchFt)
                    return true;
            }

            return false;
        }

        private static void ApplyRailingPhysicalSideBoundaries(
            Element railing,
            XYZ xP,
            XYZ yP,
            double runX0,
            double runX1,
            double referenceCenterY,
            double nominalMinY,
            double nominalMaxY,
            double sideSearchFt,
            RouteDebugLog debugLog,
            ref double clearMinY,
            ref double clearMaxY,
            out bool hasLeft,
            out bool hasRight)
        {
            hasLeft = false;
            hasRight = false;
            if (railing == null || xP == null || yP == null)
                return;

            var leftBox = new ProjectionRange2D
            {
                MinX = double.PositiveInfinity,
                MaxX = double.NegativeInfinity,
                MinY = double.PositiveInfinity,
                MaxY = double.NegativeInfinity
            };
            var rightBox = new ProjectionRange2D
            {
                MinX = double.PositiveInfinity,
                MaxX = double.NegativeInfinity,
                MinY = double.PositiveInfinity,
                MaxY = double.NegativeInfinity
            };

            int leftPointCount = 0;
            int rightPointCount = 0;
            double xToleranceFt = MmToInternal(5.0);
            var solids = new List<Solid>();
            AddElementAndDependentSolids(railing, solids);

            foreach (Solid solid in GetValidSolids(solids))
            {
                AccumulateRailingLocalBoundingPoints(
                    solid,
                    xP,
                    yP,
                    runX0,
                    runX1,
                    xToleranceFt,
                    referenceCenterY,
                    nominalMinY,
                    nominalMaxY,
                    sideSearchFt,
                    ref leftBox,
                    ref rightBox,
                    ref leftPointCount,
                    ref rightPointCount);
            }

            double minLongitudinalSpanFt = GetMinRailingPathObstacleOverlapFt(runX1 - runX0);
            double leftSpanFt = leftPointCount > 0 ? leftBox.MaxX - leftBox.MinX : 0.0;
            double rightSpanFt = rightPointCount > 0 ? rightBox.MaxX - rightBox.MinX : 0.0;
            hasLeft = leftPointCount > 0 && leftSpanFt >= minLongitudinalSpanFt;
            hasRight = rightPointCount > 0 && rightSpanFt >= minLongitudinalSpanFt;

            double clearanceFt = MmToInternal(1.0);
            if (hasLeft)
                clearMinY = Math.Max(clearMinY, leftBox.MaxY + clearanceFt);
            if (hasRight)
                clearMaxY = Math.Min(clearMaxY, rightBox.MinY - clearanceFt);

            debugLog?.Add(
                $"RunClearWidth railing clippedLocalBBox ID={IDHelper.ElIdValue(railing.Id)} " +
                $"runX={FormatFt(runX0)}..{FormatFt(runX1)} minSpan={FormatFtMm(minLongitudinalSpanFt)} " +
                $"left={(leftPointCount > 0 ? $"points={leftPointCount} span={FormatFtMm(leftSpanFt)} x={FormatFt(leftBox.MinX)}..{FormatFt(leftBox.MaxX)} y={FormatFt(leftBox.MinY)}..{FormatFt(leftBox.MaxY)} accepted={hasLeft}" : "none")} " +
                $"right={(rightPointCount > 0 ? $"points={rightPointCount} span={FormatFtMm(rightSpanFt)} x={FormatFt(rightBox.MinX)}..{FormatFt(rightBox.MaxX)} y={FormatFt(rightBox.MinY)}..{FormatFt(rightBox.MaxY)} accepted={hasRight}" : "none")} " +
                $"clear={FormatFt(clearMinY)}..{FormatFt(clearMaxY)}");
        }

        private static void AccumulateRailingLocalBoundingPoints(
            Solid solid,
            XYZ xP,
            XYZ yP,
            double runX0,
            double runX1,
            double xToleranceFt,
            double referenceCenterY,
            double nominalMinY,
            double nominalMaxY,
            double sideSearchFt,
            ref ProjectionRange2D leftBox,
            ref ProjectionRange2D rightBox,
            ref int leftPointCount,
            ref int rightPointCount)
        {
            if (solid == null || solid.Volume < 1e-9)
                return;

            double clipMinX = runX0 - xToleranceFt;
            double clipMaxX = runX1 + xToleranceFt;

            foreach (Face face in solid.Faces)
            {
                Mesh mesh;
                try { mesh = face.Triangulate(); }
                catch { continue; }

                if (mesh == null)
                    continue;

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    if (triangle == null)
                        continue;

                    var localPoints = new List<Tuple<double, double>>(9);
                    for (int j = 0; j < 3; j++)
                    {
                        XYZ p = triangle.get_Vertex(j);
                        if (p == null)
                            continue;

                        XYZ pxy = new XYZ(p.X, p.Y, 0.0);
                        localPoints.Add(Tuple.Create(pxy.DotProduct(xP), pxy.DotProduct(yP)));
                    }

                    if (localPoints.Count != 3)
                        continue;

                    AddRailingTriangleClipIntersections(localPoints, clipMinX);
                    AddRailingTriangleClipIntersections(localPoints, clipMaxX);

                    foreach (Tuple<double, double> localPoint in localPoints)
                    {
                        double tx = localPoint.Item1;
                        double ty = localPoint.Item2;
                        if (tx < clipMinX || tx > clipMaxX)
                            continue;

                        if (ty <= referenceCenterY
                            && ty > nominalMinY - sideSearchFt)
                        {
                            ExpandProjectionRange(ref leftBox, tx, ty);
                            leftPointCount++;
                        }

                        if (ty >= referenceCenterY
                            && ty < nominalMaxY + sideSearchFt)
                        {
                            ExpandProjectionRange(ref rightBox, tx, ty);
                            rightPointCount++;
                        }
                    }
                }
            }
        }

        private static void AddRailingTriangleClipIntersections(
            List<Tuple<double, double>> localPoints,
            double clipX)
        {
            if (localPoints == null || localPoints.Count < 3)
                return;

            int sourceCount = 3;
            for (int i = 0; i < sourceCount; i++)
            {
                Tuple<double, double> a = localPoints[i];
                Tuple<double, double> b = localPoints[(i + 1) % sourceCount];
                double dx = b.Item1 - a.Item1;
                if (Math.Abs(dx) < 1e-9)
                    continue;

                double t = (clipX - a.Item1) / dx;
                if (t < 0.0 || t > 1.0)
                    continue;

                double y = a.Item2 + (b.Item2 - a.Item2) * t;
                localPoints.Add(Tuple.Create(clipX, y));
            }
        }

        private static void ExpandProjectionRange(
            ref ProjectionRange2D range,
            double x,
            double y)
        {
            if (x < range.MinX) range.MinX = x;
            if (x > range.MaxX) range.MaxX = x;
            if (y < range.MinY) range.MinY = y;
            if (y > range.MaxY) range.MaxY = y;
        }

        private static bool TryGetRailingPathProjectionRanges(Element elem, XYZ xP, XYZ yP, out List<ProjectionRange2D> ranges)
        {
            ranges = new List<ProjectionRange2D>();

            Railing railing = elem as Railing;
            if (railing == null || xP == null || yP == null)
                return false;

            IEnumerable<Curve> curves = null;
            try
            {
                MethodInfo method = typeof(Railing).GetMethod("GetPath", Type.EmptyTypes);
                object raw = method == null ? null : method.Invoke(railing, null);
                curves = raw as IEnumerable<Curve>;
            }
            catch
            {
            }

            if (curves == null)
                return false;

            double pathAllowanceFt = MmToInternal(10.0);

            foreach (Curve curve in curves)
            {
                if (curve == null)
                    continue;

                var points = new List<XYZ>();
                try
                {
                    IList<XYZ> tessellated = curve.Tessellate();
                    if (tessellated != null)
                        points.AddRange(tessellated.Where(x => x != null));
                }
                catch
                {
                }

                if (points.Count == 0)
                {
                    try
                    {
                        points.Add(curve.GetEndPoint(0));
                        points.Add(curve.GetEndPoint(1));
                    }
                    catch
                    {
                    }
                }

                if (points.Count == 0)
                    continue;

                double minX = double.PositiveInfinity;
                double maxX = double.NegativeInfinity;
                double minY = double.PositiveInfinity;
                double maxY = double.NegativeInfinity;

                foreach (XYZ point in points)
                {
                    if (point == null)
                        continue;

                    XYZ p = new XYZ(point.X, point.Y, 0.0);
                    double px = p.DotProduct(xP);
                    double py = p.DotProduct(yP);

                    if (px < minX) minX = px;
                    if (px > maxX) maxX = px;
                    if (py < minY) minY = py;
                    if (py > maxY) maxY = py;
                }

                if (!double.IsInfinity(minX) && !double.IsInfinity(maxX) && !double.IsInfinity(minY) && !double.IsInfinity(maxY))
                {
                    ranges.Add(new ProjectionRange2D
                    {
                        MinX = minX - pathAllowanceFt,
                        MaxX = maxX + pathAllowanceFt,
                        MinY = minY - pathAllowanceFt,
                        MaxY = maxY + pathAllowanceFt
                    });
                }
            }

            return ranges.Count > 0;
        }

        private static Solid TryCreateRunClearWidthObstacleTestSolid(StairsRun run, XYZ bottom, XYZ top, XYZ yP, double widthFt, RouteDebugLog debugLog)
        {
            if (run == null || bottom == null || top == null || yP == null || yP.GetLength() < 1e-9 || widthFt <= 1e-9)
                return null;

            if (!TryGetElementSolidsZRange(run, out double minZ, out double maxZ))
            {
                BoundingBoxXYZ box = null;
                try { box = run.get_BoundingBox(null); } catch { }
                if (box == null)
                    return null;

                minZ = box.Min.Z;
                maxZ = box.Max.Z;
            }

            minZ -= MmToInternal(150.0);
            maxZ += MmToInternal(1400.0);
            if (maxZ - minZ < MmToInternal(300.0))
                maxZ = minZ + MmToInternal(300.0);

            yP = yP.Normalize();
            XYZ halfW = yP * (widthFt / 2.0);
            XYZ bottomCenter = new XYZ(bottom.X, bottom.Y, minZ);
            XYZ topCenter = new XYZ(top.X, top.Y, minZ);

            XYZ SL = bottomCenter - halfW;
            XYZ SR = bottomCenter + halfW;
            XYZ ER = topCenter + halfW;
            XYZ EL = topCenter - halfW;
            XYZ up = XYZ.BasisZ * (maxZ - minZ);

            Solid solid = BuildPrismFrom8Points(SL, SR, ER, EL, SL + up, SR + up, ER + up, EL + up);
            if (solid == null || solid.Volume < 1e-9)
                return null;

            debugLog?.Add($"RunClearWidth obstacle test solid z={FormatFtMm(minZ)}..{FormatFtMm(maxZ)} width={FormatFtMm(widthFt)}");
            return solid;
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

        private static bool TryGetElementSolidsZRange(Element elem, out double minZ, out double maxZ)
        {
            minZ = double.PositiveInfinity;
            maxZ = double.NegativeInfinity;

            if (elem == null)
                return false;

            var solids = new List<Solid>();
            AddElementSolids(elem, solids);
            foreach (Solid solid in GetValidSolids(solids))
            {
                if (!TryGetSolidZRange(solid, out double solidMinZ, out double solidMaxZ))
                    continue;

                if (solidMinZ < minZ) minZ = solidMinZ;
                if (solidMaxZ > maxZ) maxZ = solidMaxZ;
            }

            return maxZ > minZ;
        }

        private static bool TryGetSolidZRange(Solid solid, out double minZ, out double maxZ)
        {
            minZ = double.PositiveInfinity;
            maxZ = double.NegativeInfinity;

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

                        if (p.Z < minZ) minZ = p.Z;
                        if (p.Z > maxZ) maxZ = p.Z;
                        hasPoints = true;
                    }
                }
            }

            return hasPoints;
        }

        private static bool TryGetClearWidthObstacleProjectionRange(
            Solid obstacleTestSolid,
            Element obstacleElement,
            Solid obstacleSolid,
            XYZ xP,
            XYZ yP,
            out double minX,
            out double maxX,
            out double minY,
            out double maxY)
        {
            minX = double.PositiveInfinity;
            maxX = double.NegativeInfinity;
            minY = double.PositiveInfinity;
            maxY = double.NegativeInfinity;

            if (obstacleSolid == null || obstacleSolid.Volume < 1e-9)
                return false;

            Solid rangeSolid = obstacleSolid;
            if (obstacleTestSolid != null && obstacleTestSolid.Volume > 1e-9)
            {
                try
                {
                    Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(obstacleTestSolid, obstacleSolid, BooleanOperationsType.Intersect);
                    if (intersection == null || intersection.Volume <= GetMinIntersectionVolumeFt3())
                        return false;

                    if (!HasMinimumIntersectionThickness(intersection, GetClearWidthObstacleMinIntersectionThicknessFt(obstacleElement)))
                        return false;

                    rangeSolid = intersection;
                }
                catch
                {
                    return false;
                }
            }

            return TryGetSolidProjectionRange(rangeSolid, xP, yP, out minX, out maxX, out minY, out maxY);
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

        private static bool TryGetRunGeometryWidthRange(StairsRun run, XYZ xP, XYZ yP, XYZ bottom, XYZ top, double lenPlan, out double minY, out double maxY)
        {
            minY = double.PositiveInfinity;
            maxY = double.NegativeInfinity;

            if (run == null || xP == null || yP == null)
                return false;

            double runX0 = new XYZ(bottom.X, bottom.Y, 0).DotProduct(xP);
            double runX1 = new XYZ(top.X, top.Y, 0).DotProduct(xP);
            if (runX1 < runX0)
            {
                double tmp = runX0;
                runX0 = runX1;
                runX1 = tmp;
            }

            double xPaddingFt = Math.Max(MmToInternal(100.0), lenPlan * 0.02);
            bool hasRange = false;

            var solids = new List<Solid>();
            AddElementSolids(run, solids);
            foreach (Solid solid in GetValidSolids(solids))
            {
                if (!TryGetSolidProjectionRange(solid, xP, yP, out double solidMinX, out double solidMaxX, out double solidMinY, out double solidMaxY))
                    continue;

                if (solidMaxX < runX0 - xPaddingFt || solidMinX > runX1 + xPaddingFt)
                    continue;

                if (solidMinY < minY) minY = solidMinY;
                if (solidMaxY > maxY) maxY = solidMaxY;
                hasRange = true;
            }

            return hasRange && maxY - minY > MmToInternal(50.0);
        }

        private static double GetRunWidthSideAllowanceFt()
        {
            return 0.0;
        }

        private static double GetRunWidthFt(StairsRun run, XYZ bottom, XYZ top, XYZ xPlan, XYZ yDir)
        {
            if (TryGetActualRunWidthFt(run, out double actualWidthFt))
                return actualWidthFt;

            if (xPlan == null || xPlan.GetLength() < 1e-9 || yDir == null || yDir.GetLength() < 1e-9)
            {
#if Debug2023 || Debug2024 || Revit2023 || Revit2024
                return UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);
#else
                return UnitUtils.ConvertToInternalUnits(1000.0, DisplayUnitType.DUT_MILLIMETERS);
#endif
            }

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

        private static bool TryGetActualRunWidthFt(StairsRun run, out double widthFt)
        {
            widthFt = 0.0;

            if (run == null)
                return false;

            try
            {
                var prop = run.GetType().GetProperty("ActualRunWidth");
                if (prop == null)
                    return false;

                object v = prop.GetValue(run, null);
                if (v is double d && d > 1e-9)
                {
                    widthFt = d;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }
    }
}