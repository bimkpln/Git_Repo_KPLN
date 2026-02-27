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

        // Класс выбора лестницы в Revit
        private sealed class StairsSelectionFilter : ISelectionFilter
        {
            private readonly Document _doc;
            public StairsSelectionFilter(Document doc) { _doc = doc; }

            public bool AllowElement(Element elem)
            {
                return elem is Stairs;
            }

            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        // Построение путей эвакуации
        private void CreateEvacuationRoutes()
        {
            var stairs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().OfType<Stairs>().ToList();
            int stairsCount = stairs.Count;

            var dlg = new EvacuationRoutesDialog(stairsCount);
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
                {
                    return;
                }

                targetStairs = new List<Stairs> { picked };
            }

            using (var t = new Transaction(doc, "KPLN: Построение путей эвакуации"))
            {
                t.Start();


                int totalRunsInDoc = stairs.Sum(x => (x.GetStairsRuns()?.Count ?? 0));
                int totalRunsInScope = targetStairs.Sum(x => (x.GetStairsRuns()?.Count ?? 0));

                int createdRuns = 0; 
                int createdStairs = 0;
                var failedStairsIds = new List<int>(); 
                var failedRunIds = new List<int>(); 

                foreach (var s in targetStairs)
                {
                    int stairCreatedRuns;
                    List<int> stairFailedRuns;

                    bool okStair = TryCreateRouteBodyOnStair(doc, s, data, out stairCreatedRuns, out stairFailedRuns);

                    if (okStair)
                    {
                        createdStairs++;
                        createdRuns += stairCreatedRuns;
                    }
                    else
                    {
                        failedStairsIds.Add(s.Id.IntegerValue);
                    }

                    if (stairFailedRuns != null && stairFailedRuns.Count > 0)
                        failedRunIds.AddRange(stairFailedRuns);
                }



                t.Commit();

                string scopeInfo = (targetStairs.Count == 1) ? $"Режим: выбрать лестницу (ID: {targetStairs[0].Id.IntegerValue})" : "Режим: все лестницы";

                string failedStairsText = failedStairsIds.Count == 0 ? "Все лестницы обработаны" : $"Необработанные лестницы (ID): {string.Join(", ", failedStairsIds.Distinct().OrderBy(x => x))}";
                string failedRunsText = failedRunIds.Count == 0 ? "Все лестничные марши обработаны" : $"Необработанные марши (ID): {string.Join(", ", failedRunIds.Distinct().OrderBy(x => x))}";

                TaskDialog.Show("Готово", $"{scopeInfo}\n" +
                    $"Лестниц [маршей] в документе: {stairsCount} [{totalRunsInDoc}]\n" +
                    $"Лестниц [маршей] обработано: {createdStairs} [{createdRuns}]\n\n" +
                    $"{failedStairsText}\n" +
                    $"{failedRunsText}" 
                );
            }
        }

        // Выбор конкретной лестницы в Revit
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





















        // Построение путей эвакуации на лестничных маршах и лестницах
        private bool TryCreateRouteBodyOnStair(Document doc, Stairs stairs, EvacuationRoutesDialogResult data, out int createdRuns, out List<int> failedRunIds)
        {
            createdRuns = 0;
            failedRunIds = new List<int>();

            var runIds = stairs.GetStairsRuns();
            if (runIds == null || runIds.Count == 0) return false;

            double heightFt = UnitUtils.ConvertToInternalUnits(data.HeightMm, UnitTypeId.Millimeters);
            if (heightFt <= 1e-9) return false;

            double epsMm = 1.0;
            double epsFt = UnitUtils.ConvertToInternalUnits(epsMm, UnitTypeId.Millimeters);

            foreach (ElementId runId in runIds)
            {
                bool runCreated = false;

                StairsRun run = doc.GetElement(runId) as StairsRun;
                if (run == null) { failedRunIds.Add(runId.IntegerValue); continue; }

                CurveLoop path;
                try { path = run.GetStairsPath(); }
                catch { failedRunIds.Add(runId.IntegerValue); continue; }

                if (path == null) { failedRunIds.Add(runId.IntegerValue); continue; }
                var curves = path.ToList();
                if (curves.Count == 0) { failedRunIds.Add(runId.IntegerValue); continue; }

                XYZ p0 = curves.First().GetEndPoint(0);
                XYZ p1 = curves.Last().GetEndPoint(1);

                XYZ bottomCenter, topCenter;
                if (p0.Z <= p1.Z) { bottomCenter = p0; topCenter = p1; }
                else { bottomCenter = p1; topCenter = p0; }

                double widthFt = data.UseRunWidth
                    ? GetRunWidthFt(run, bottomCenter, topCenter)
                    : UnitUtils.ConvertToInternalUnits(data.WidthMm, UnitTypeId.Millimeters);

                if (widthFt <= 1e-9) { failedRunIds.Add(runId.IntegerValue); continue; }

                XYZ xP = new XYZ(topCenter.X - bottomCenter.X, topCenter.Y - bottomCenter.Y, 0.0);
                double lenPlan = xP.GetLength();
                if (lenPlan < 1e-9) { failedRunIds.Add(runId.IntegerValue); continue; }
                xP = xP.Normalize();

                XYZ yP = XYZ.BasisZ.CrossProduct(xP);
                if (yP.GetLength() < 1e-9) yP = XYZ.BasisY;
                yP = yP.Normalize();

                XYZ halfW = yP * (widthFt / 2.0);

                Plane undersidePlane;
                if (!TryGetBestUndersidePlane(run, out undersidePlane))
                {
                    failedRunIds.Add(runId.IntegerValue);
                    continue;
                }

                double midGapFt = GetMidTreadGapFt_NoDeps(run, undersidePlane, bottomCenter, xP, lenPlan, widthFt, yP);
                if (midGapFt <= 1e-9) midGapFt = 0.0;

                double liftFt = midGapFt + epsFt;

                XYZ SL_xy = new XYZ(bottomCenter.X, bottomCenter.Y, 0) - halfW;
                XYZ SR_xy = new XYZ(bottomCenter.X, bottomCenter.Y, 0) + halfW;
                XYZ EL_xy = new XYZ(topCenter.X, topCenter.Y, 0) - halfW;
                XYZ ER_xy = new XYZ(topCenter.X, topCenter.Y, 0) + halfW;

                double zSL = GetPlaneZAtXY(undersidePlane, SL_xy.X, SL_xy.Y);
                double zSR = GetPlaneZAtXY(undersidePlane, SR_xy.X, SR_xy.Y);
                double zEL = GetPlaneZAtXY(undersidePlane, EL_xy.X, EL_xy.Y);
                double zER = GetPlaneZAtXY(undersidePlane, ER_xy.X, ER_xy.Y);

                XYZ SL = new XYZ(SL_xy.X, SL_xy.Y, zSL + liftFt);
                XYZ SR = new XYZ(SR_xy.X, SR_xy.Y, zSR + liftFt);
                XYZ ER = new XYZ(ER_xy.X, ER_xy.Y, zER + liftFt);
                XYZ EL = new XYZ(EL_xy.X, EL_xy.Y, zEL + liftFt);

                XYZ up = XYZ.BasisZ * heightFt;
                XYZ SLt = SL + up;
                XYZ SRt = SR + up;
                XYZ ERt = ER + up;
                XYZ ELt = EL + up;

                Solid solid = BuildPrismFrom8Points(SL, SR, ER, EL, SLt, SRt, ERt, ELt);
                if (solid == null || solid.Volume < 1e-9)
                {
                    failedRunIds.Add(runId.IntegerValue);
                    continue;
                }

                ElementId catId = new ElementId(BuiltInCategory.OST_Site);
                DirectShape ds = DirectShape.CreateElement(doc, catId);

                ds.ApplicationId = "KPLN_Tools";
                ds.ApplicationDataId = run.Id.IntegerValue.ToString();
                ds.Name = $"ПЭ_{stairs.Id.IntegerValue}{run.Id.IntegerValue}";

                ds.SetShape(new List<GeometryObject> { solid });

                runCreated = true;

                if (runCreated)
                    createdRuns++;
                else
                    failedRunIds.Add(runId.IntegerValue);
            }

            return createdRuns > 0;
        }

        private static bool TryGetTopZByVerticalIntersect(StairsRun run, XYZ pointXY, out double zTop)
        {
            zTop = double.NegativeInfinity;

            Options opt = new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = false,
                IncludeNonVisibleObjects = false
            };

            GeometryElement ge = run.get_Geometry(opt);
            if (ge == null) return false;

            var solids = new List<Solid>();
            CollectSolidsRecursive(ge, Transform.Identity, solids);
            if (solids.Count == 0) return false;

            double big = UnitUtils.ConvertToInternalUnits(20000.0, UnitTypeId.Millimeters);
            XYZ pTop = new XYZ(pointXY.X, pointXY.Y, big);
            XYZ pBot = new XYZ(pointXY.X, pointXY.Y, -big);
            Line line = Line.CreateBound(pTop, pBot);

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

                    if (a != null && a.Z > zTop) zTop = a.Z;
                    if (b != null && b.Z > zTop) zTop = b.Z;
                }
            }

            return !double.IsNegativeInfinity(zTop);
        }

        private static double GetMidTreadGapFt_NoDeps(StairsRun run, Plane undersidePlane, XYZ bottomCenter, XYZ xP, double lenPlan, double widthFt, XYZ yP)
        {
            double[] u = new double[] { 0.45, 0.50, 0.55 };

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

                    if (!TryGetTopZByVerticalIntersect(run, p, out double zTop))
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

        private static bool TryGetBestUndersidePlane(StairsRun run, out Plane plane)
        {
            plane = null;

            Options opt = new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = false,
                IncludeNonVisibleObjects = false
            };

            GeometryElement ge = run.get_Geometry(opt);
            if (ge == null) return false;

            PlanarFace best = null;
            double bestScore = double.NegativeInfinity;

            foreach (GeometryObject go in ge)
            {
                Solid s = go as Solid;
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
                tsb.AddFace(new TessellatedFace(
                    new List<XYZ> { a, b, c }, ElementId.InvalidElementId));

                tsb.AddFace(new TessellatedFace(
                    new List<XYZ> { a, c, d }, ElementId.InvalidElementId));
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
                return UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);

            XYZ xPlan = dirPlan.Normalize();
            XYZ yDir = XYZ.BasisZ.CrossProduct(xPlan);
            if (yDir.GetLength() < 1e-9) yDir = XYZ.BasisY;
            yDir = yDir.Normalize();

            BoundingBoxXYZ bb = run.get_BoundingBox(null);
            if (bb == null)
                return UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);

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
                w = UnitUtils.ConvertToInternalUnits(1000.0, UnitTypeId.Millimeters);

            return w;
        }
    }
}