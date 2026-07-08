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

            public bool AllowElement(Element elem) => elem is Stairs;
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
            var stairs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().OfType<Stairs>().ToList();

            int stairsCount = stairs.Count;
            var evacuationWorksets = GetEvacuationWorksetOptions(doc);

            var dlg = new EvacuationRoutesDialog(stairsCount, evacuationWorksets);
            new WindowInteropHelper(dlg) { Owner = uiapp.MainWindowHandle };

            bool? ok = dlg.ShowDialog();
            if (ok != true || dlg.Result == null)
                return;

            var data = dlg.Result;

            List<Stairs> targetStairs = stairs;

            if (data.PickSingleStair)
            {
                try { dlg.Hide(); } catch { }

                Stairs picked = PickSingleStairs(uiapp, doc);
                if (picked == null)
                    return;

                targetStairs = new List<Stairs> { picked };
            }

            using (var t = new Transaction(doc, "KPLN: Построение путей эвакуации"))
            {
                t.Start();

                int totalRunsInDoc = stairs.Sum(x => (x.GetStairsRuns()?.Count ?? 0));
                int totalLandingsInDoc = stairs.Sum(x => (x.GetStairsLandings()?.Count ?? 0));

                int createdRuns = 0;
                int createdStairs = 0;
                int createdLandings = 0;

                var failedStairsIds = new List<int>();
                var failedRunIds = new List<int>();
                var failedLandingIds = new List<int>();

                foreach (var s in targetStairs)
                {
                    int stairCreatedRuns;
                    int stairCreatedLandings;
                    List<int> stairFailedRuns;
                    List<int> stairFailedLandings;

                    bool okStair = TryCreateRouteBodyOnStair(
                        doc, s, data,
                        out stairCreatedRuns, out stairCreatedLandings,
                        out stairFailedRuns, out stairFailedLandings);

                    if (okStair)
                    {
                        createdStairs++;
                        createdRuns += stairCreatedRuns;
                        createdLandings += stairCreatedLandings;
                    }
                    else
                    {
                        failedStairsIds.Add(IDHelper.ElIdInt(s.Id));
                    }

                    if (stairFailedRuns != null && stairFailedRuns.Count > 0)
                        failedRunIds.AddRange(stairFailedRuns);

                    if (stairFailedLandings != null && stairFailedLandings.Count > 0)
                        failedLandingIds.AddRange(stairFailedLandings);
                }

                t.Commit();

                bool isSingle = targetStairs.Count == 1;

                // Приводим списки к нормальному виду
                var stairsFail = failedStairsIds.Distinct().OrderBy(x => x).ToList();
                var runsFail = failedRunIds.Distinct().OrderBy(x => x).ToList();
                var landingsFail = failedLandingIds.Distinct().OrderBy(x => x).ToList();

                if (!isSingle)
                {
                    // ===== РЕЖИМ: ВСЕ ЛЕСТНИЦЫ =====
                    string msg =
                        "Операция завершена.\n\n" +
                        FormatIdsLine("Необработанные лестницы", stairsFail) + "\n" +
                        FormatIdsLine("Необработанные марши", runsFail);

                    bool needLog = TooManyIdsForDialog(stairsFail.Concat(runsFail));

                    if (!needLog)
                    {
                        TaskDialog.Show("Готово", msg);
                    }
                    else
                    {
                        TaskDialog td = new TaskDialog("Готово");
                        td.MainInstruction = "Операция завершена.";
                        td.MainContent =
                            "Список необработанных ID слишком большой для удобного просмотра.\n" +
                            "Можно сохранить TXT-лог на рабочий стол.\n\n" +
                            FormatIdsLine("Необработанные лестницы", stairsFail) + "\n" +
                            FormatIdsLine("Необработанные марши", runsFail);

                        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Сохранить TXT-лог на рабочий стол");
                        td.CommonButtons = TaskDialogCommonButtons.Close;

                        TaskDialogResult res = td.Show();
                        if (res == TaskDialogResult.CommandLink1)
                        {
                            try
                            {
                                string path = SaveFailuresLogToDesktop(stairsFail, runsFail, landingsFail);
                                TaskDialog.Show("Лог сохранён", $"TXT-лог сохранён:\n{path}");
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Ошибка", $"Не удалось сохранить лог.\n\n{ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    // ===== РЕЖИМ: ОДНА ЛЕСТНИЦА =====
                    int stairId = IDHelper.ElIdInt(targetStairs[0].Id);

                    bool stairFailed = stairsFail.Contains(stairId);
                    bool anyRunsFailed = runsFail.Count > 0;
                    bool anyLandingsFailed = landingsFail.Count > 0;

                    if (!stairFailed && !anyRunsFailed && !anyLandingsFailed)
                    {
                        TaskDialog.Show("Готово", $"Операция завершена.\nВсе элементы обработаны.\n\nЛестница ID: {stairId}");
                    }
                    else
                    {
                        var lines = new List<string>
                        {
                            "Операция завершена.",
                            $"Лестница ID: {stairId}",
                            ""
                        };

                        if (stairFailed) lines.Add("Не обработана сама лестница.");
                        if (anyRunsFailed) lines.Add(FormatIdsLine("Необработанные марши", runsFail));
                        if (anyLandingsFailed) lines.Add(FormatIdsLine("Необработанные площадки", landingsFail));

                        TaskDialog.Show("Готово", string.Join("\n", lines));
                    }
                }
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

            System.IO.File.WriteAllLines(path, lines);
            return path;
        }


        private static Stairs PickSingleStairs(UIApplication uiapp, Document doc)
        {
            try
            {
                var uidoc = uiapp.ActiveUIDocument;

                Reference r = uidoc.Selection.PickObject(ObjectType.Element, new StairsSelectionFilter(doc), "Выберите одну лестницу (Esc — Отмена)");

                if (r == null) return null;

                return doc.GetElement(r.ElementId) as Stairs;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        // =========================
        // ЛЕСТНИЦА: МАРШИ+ПЛОЩАДКИ
        // =========================

        private bool TryCreateRouteBodyOnStair(Document doc, Stairs stairs, EvacuationRoutesDialogResult data,
            out int createdRuns, out int createdLandings, out List<int> failedRunIds, out List<int> failedLandingIds)
        {
            createdRuns = 0;
            createdLandings = 0;
            failedRunIds = new List<int>();
            failedLandingIds = new List<int>();

            var runIds = stairs.GetStairsRuns();
            var landingIds = stairs.GetStairsLandings();

            bool hasRuns = runIds != null && runIds.Count > 0;
            bool hasLandings = landingIds != null && landingIds.Count > 0;

            if (!hasRuns && !hasLandings) return false;


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
                    bool okRun = TryCreateRouteBodyOnRun(doc, stairs, run, data, heightFt, epsFt, out info);

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

                    bool okLanding = TryCreateRouteBodyOnLanding(doc, stairs, landing, runs, runInfos, data, heightFt);
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

        // Создаёт или обновляет DirectShape с заданным appId/appDataId. Возвращает сам элемент (на будущее, если надо логировать).
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

        // =========================
        // МАРШ
        // =========================
        private bool TryCreateRouteBodyOnRun(Document doc, Stairs stairs, StairsRun run, EvacuationRoutesDialogResult data, double heightFt, double epsFt, out RunRouteBodyInfo runInfo)
        {
            runInfo = null;

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
            }
            else
            {
                widthFt = MmToInternal(data.WidthMm);
            }

            if (widthFt <= 1e-9) return false;

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

            XYZ up = XYZ.BasisZ * heightFt;
            XYZ SLt = SL + up;
            XYZ SRt = SR + up;
            XYZ ERt = ER + up;
            XYZ ELt = EL + up;

            Solid solid = BuildPrismFrom8Points(SL, SR, ER, EL, SLt, SRt, ERt, ELt);
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

            // СОЗДАТЬ ИЛИ ОБНОВИТЬ
            UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", IDHelper.ElIdValue(run.Id).ToString(), $"ПЭ_{IDHelper.ElIdValue(stairs.Id)}{IDHelper.ElIdValue(run.Id)}", solid, data);
            return true;
        }

        // =========================
        //   ПЛОЩАДКА: КОРОБКА ПО ЭКСТРЕМУМАМ УГЛОВ ДВУХ МАРШЕЙ
        // =========================
        private bool TryCreateRouteBodyOnLanding(Document doc, Stairs stairs, StairsLanding landing, List<StairsRun> runsInSameStair, List<RunRouteBodyInfo> runInfos, EvacuationRoutesDialogResult data, double heightFt)
        {
            if (runInfos == null || runInfos.Count < 2)
                return false;

            BoundingBoxXYZ bbL = landing.get_BoundingBox(null);
            if (bbL == null) return false;

            XYZ landingCenter = new XYZ(
                (bbL.Min.X + bbL.Max.X) * 0.5,
                (bbL.Min.Y + bbL.Max.Y) * 0.5,
                (bbL.Min.Z + bbL.Max.Z) * 0.5);

            // Для каждого марша выбираем ближайший к площадке торец (Top/Bottom)
            var candidates = new List<(RunRouteBodyInfo run, EndFace face, XYZ faceCenter, double dist)>();
            foreach (var ri in runInfos)
            {
                double dTop = new XYZ(ri.TopEnd.Center.X - landingCenter.X, ri.TopEnd.Center.Y - landingCenter.Y, 0).GetLength();
                double dBot = new XYZ(ri.BottomEnd.Center.X - landingCenter.X, ri.BottomEnd.Center.Y - landingCenter.Y, 0).GetLength();

                bool useTop = dTop <= dBot;
                EndFace f = useTop ? ri.TopEnd : ri.BottomEnd;

                candidates.Add((ri, f, f.Center, useTop ? dTop : dBot));
            }

            if (candidates.Count < 2)
                return false;

            // Берём 2 ближайших к площадке марша/торца
            var two = candidates.OrderBy(x => x.dist).Take(2).ToList();
            var A = two[0];
            var B = two[1];

            // xDir — направление "длины" блока на площадке: между центрами выбранных торцов
            XYZ xDir = new XYZ(B.faceCenter.X - A.faceCenter.X, B.faceCenter.Y - A.faceCenter.Y, 0);
            if (xDir.GetLength() < 1e-9)
                xDir = new XYZ(A.run.XDirPlan.X, A.run.XDirPlan.Y, 0);
            if (xDir.GetLength() < 1e-9)
                return false;
            xDir = xDir.Normalize();

            // yDir — направление "глубины" (красная линия) поперёк марша. Берём поперечную ось одного марша и ортогонализируем относительно xDir
            XYZ yBase = new XYZ(A.run.YDirPlan.X, A.run.YDirPlan.Y, 0);
            if (yBase.GetLength() < 1e-9) yBase = new XYZ(B.run.YDirPlan.X, B.run.YDirPlan.Y, 0);
            if (yBase.GetLength() < 1e-9) yBase = XYZ.BasisZ.CrossProduct(xDir);

            yBase = yBase.Normalize();
            XYZ yDir = yBase - xDir * (yBase.DotProduct(xDir));
            if (yDir.GetLength() < 1e-9)
                yDir = XYZ.BasisZ.CrossProduct(xDir);

            yDir = new XYZ(yDir.X, yDir.Y, 0);
            if (yDir.GetLength() < 1e-9) return false;
            yDir = yDir.Normalize();

            // X-границы: строго min/max по углам двух блоков маршей (касание по граням)
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

            // SPAN по Y из углов блоков (нужно только чтобы понять "где грань" и "какая сторона площадки")
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

            if (landingCY > spanMaxY + tol)
            {
                // Площадка по + стороне: начинаем от грани spanMaxY и уходим в сторону площадки на depth
                minY = spanMaxY;
                maxY = spanMaxY + depthFt;
            }
            else if (landingCY < spanMinY - tol)
            {
                // Площадка по - стороне: начинаем от грани spanMinY и уходим в сторону площадки на depth
                maxY = spanMinY;
                minY = spanMinY - depthFt;
            }
            else
            {
                // Площадка "между" гранями по yDir — выбираем ближайшую грань к центру площадки и уходим к площадке
                double distToMin = Math.Abs(landingCY - spanMinY);
                double distToMax = Math.Abs(spanMaxY - landingCY);

                if (distToMax <= distToMin)
                {
                    minY = spanMaxY;
                    maxY = spanMaxY + depthFt;
                }
                else
                {
                    maxY = spanMinY;
                    minY = spanMinY - depthFt;
                }
            }

            // Z-низ/высота: - старт от "нижнего блока" и высота = height блока
            double baseZ = Math.Min(A.face.MinZBottom, B.face.MinZBottom);
            double h = heightFt;
            if (h <= 1e-9) return false;

            XYZ P1 = xDir * minX + yDir * minY + XYZ.BasisZ * baseZ;
            XYZ P2 = xDir * maxX + yDir * minY + XYZ.BasisZ * baseZ;
            XYZ P3 = xDir * maxX + yDir * maxY + XYZ.BasisZ * baseZ;
            XYZ P4 = xDir * minX + yDir * maxY + XYZ.BasisZ * baseZ;

            XYZ up = XYZ.BasisZ * h;

            Solid solid = BuildPrismFrom8Points(P1, P2, P3, P4, P1 + up, P2 + up, P3 + up, P4 + up);
            if (solid == null || solid.Volume < 1e-9)
                return false;

            // СОЗДАТЬ ИЛИ ОБНОВИТЬ
            UpsertRouteShape(doc, new ElementId(BuiltInCategory.OST_Site), "KPLN_Tools", IDHelper.ElIdValue(landing.Id).ToString(), $"ПЭ_Л_{IDHelper.ElIdValue(stairs.Id)}_{IDHelper.ElIdValue(landing.Id)}", solid, data);
            return true;
        }

        // =========================
        //   НИЗ / ЗАЗОРЫ
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

            // Отделка ступеней иногда живёт внутри родительской Stairs, а не в StairsRun.
            // Используем её только как небольшую добавку над найденным верхом текущего марша.
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
        //   ПОСТРОЕНИЕ SOLID
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
        //   ШИРИНА МАРША
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