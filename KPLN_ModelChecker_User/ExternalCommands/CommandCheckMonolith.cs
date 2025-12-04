using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ModelChecker_User.Forms;


namespace KPLN_ModelChecker_User.ExternalCommands
{
    public class SkippedElementInfo
    {
        public Element Element { get; }
        public string Origin { get; }

        public SkippedElementInfo(Element element, string origin)
        {
            Element = element;
            Origin = origin;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandCheckMonolith : IExternalCommand
    {
        internal const string PluginName = "АР/КР. Проверка на совпадение монолита";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var checkMonolithSettingsWindow = new Forms.CheckMonolithSettings(uiApp, doc);

            if (checkMonolithSettingsWindow.ShowDialog() == true)
            {
                List<string> categories = checkMonolithSettingsWindow.SelectedCategories.ToList();
                double toleranceValue = checkMonolithSettingsWindow.Tolerance;

                List<Element> allSelectedElements = new List<Element>();
                List<(Element element, Transform transform)> allSelectedLinkElements = new List<(Element, Transform)>();

                List<Solid> allSelectedElementsSolidsDiffusion = new List<Solid>();
                List<Solid> allSelectedLinkElementsSolidsDiffusion = new List<Solid>();

                try
                {
                    // Выбор категорий
                    List<BuiltInCategory> categoryList = new List<BuiltInCategory>();
                    if (categories.Contains("Стены"))
                    {
                        categoryList.Add(BuiltInCategory.OST_Walls);
                        categoryList.Add(BuiltInCategory.OST_StructuralColumns);
                        categoryList.Add(BuiltInCategory.OST_StructuralFraming);
                    }
                    if (categories.Contains("Перекрытия"))
                    {
                        categoryList.Add(BuiltInCategory.OST_Floors);
                    }
                    if (categories.Contains("Лестницы"))
                    {
                        categoryList.Add(BuiltInCategory.OST_Stairs);
                        categoryList.Add(BuiltInCategory.OST_StairsRailing);
                        categoryList.Add(BuiltInCategory.OST_StairsLandings);
                        categoryList.Add(BuiltInCategory.OST_StairsRuns);
                        categoryList.Add(BuiltInCategory.OST_Floors);
                    }
                    BuiltInCategory[] allowedCategories = categoryList.Distinct().ToArray();

                    // Убираем ненужное
                    Func<Element, bool> IsElementAllowed = (e) =>
                    {
                        if (e == null) return false;

                        Category cat = e.Category;
                        if (cat == null || !allowedCategories.Contains((BuiltInCategory)cat.Id.IntegerValue))
                            return false;

                        string name = e.Name?.ToLowerInvariant() ?? "";
                        if (name.Contains("проем") || name.Contains("врезанные профили"))
                            return false;

                        return true;
                    };

                    IList<Reference> currentRefs = uiDoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Выделите рамкой элементы в текущей модели");

                    foreach (Reference r in currentRefs)
                    {
                        Element el = doc.GetElement(r.ElementId);
                        if (IsElementAllowed(el))
                        {
                            allSelectedElements.Add(el);
                        }
                    }

                    IList<Reference> linkedRefs = uiDoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        "Выделите рамкой элементы в связанных моделях");

                    foreach (Reference r in linkedRefs)
                    {
                        if (r.LinkedElementId != ElementId.InvalidElementId)
                        {
                            RevitLinkInstance linkInstance = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            if (linkInstance != null)
                            {
                                Document linkedDoc = linkInstance.GetLinkDocument();
                                Element linkedElement = linkedDoc.GetElement(r.LinkedElementId);
                                if (IsElementAllowed(linkedElement))
                                {
                                    Transform transform = linkInstance.GetTransform();
                                    allSelectedLinkElements.Add((linkedElement, transform));
                                }
                            }
                        }
                    }

                    if (allSelectedElements.Count == 0 || allSelectedLinkElements.Count == 0)
                    {
                        TaskDialog.Show("Ошибка", "Не выбраны подходящие элементы");
                        return Result.Failed;
                    }

                    // Создание общего Solid
                    List<Element> skippedMainElements;
                    List<Element> skippedLinkElements;
                    Solid elGrouoUnionMain = GetUnionSolid(allSelectedElements, out skippedMainElements);
                    Solid elGrouoUnionOther = GetUnionSolid(allSelectedLinkElements, out skippedLinkElements);

                    if (elGrouoUnionMain == null || elGrouoUnionOther == null)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось считать геометрию");
                        return Result.Failed;
                    }

                    string skippedMain = skippedMainElements.Any()
                        ? string.Join("\n", skippedMainElements.Select(e => $"Main ⮕ ID: {e.Id}, Имя: {e.Name}"))
                        : "Main ⮕ нет пропущенных элементов.";

                    string skippedLink = skippedLinkElements.Any()
                        ? string.Join("\n", skippedLinkElements.Select(e => $"Link ⮕ ID: {e.Id}, Имя: {e.Name}"))
                        : "Link ⮕ нет пропущенных элементов.";

                    if (skippedMainElements.Count != 0 || skippedLinkElements.Count != 0)
                    {
                        TaskDialog.Show("Предупреждение", $"Не все элементы попали в проверку:\n{skippedMain}\n{skippedLink}");
                    }

                    // Вычетание геометрий
                    Solid mainMinusOther = null;
                    Solid otherMinusMain = null;

                    try
                    {
                        mainMinusOther = BooleanOperationsUtils.ExecuteBooleanOperation(elGrouoUnionMain, elGrouoUnionOther,
                            BooleanOperationsType.Difference);
                    }
                    catch (Exception ex)
                    {
                        string info = $"elGrouoUnionMain: Volume={elGrouoUnionMain?.Volume:F2}, Faces={elGrouoUnionMain?.Faces?.Size}\n" +
                                      $"elGrouoUnionOther: Volume={elGrouoUnionOther?.Volume:F2}, Faces={elGrouoUnionOther?.Faces?.Size}\n\n" +
                                      "Пропущенные элементы:\n" +
                                      $"{skippedMain}\n" +
                                      $"{skippedLink}\n\n" +
                                      $"{ex.Message}";

                        TaskDialog.Show("Ошибка Boolean Difference", info);
                        return Result.Failed;
                    }
                    try
                    {
                        otherMinusMain = BooleanOperationsUtils.ExecuteBooleanOperation(elGrouoUnionOther, elGrouoUnionMain,
                            BooleanOperationsType.Difference);
                    }
                    catch (Exception ex)
                    {
                        string info = $"elGrouoUnionMain: Volume={elGrouoUnionMain?.Volume:F2}, Faces={elGrouoUnionMain?.Faces?.Size}\n" +
                                      $"elGrouoUnionOther: Volume={elGrouoUnionOther?.Volume:F2}, Faces={elGrouoUnionOther?.Faces?.Size}\n\n" +
                                      "Пропущенные элементы:\n" +
                                      $"{skippedMain}\n" +
                                      $"{skippedLink}\n\n" +
                                      $"{ex.Message}";

                        TaskDialog.Show("Ошибка Boolean Difference", info);
                        return Result.Failed;
                    }

                    if (mainMinusOther != null)
                    {
                        allSelectedElementsSolidsDiffusion.AddRange(SolidUtils.SplitVolumes(mainMinusOther));
                    }
                    else
                    {
                        TaskDialog.Show("Ошибка", "Не удалось обработать геометрию основной модели");
                        return Result.Failed;
                    }
                    if (otherMinusMain != null)
                    {
                        allSelectedLinkElementsSolidsDiffusion.AddRange(SolidUtils.SplitVolumes(otherMinusMain));
                    }
                    else
                    {
                        TaskDialog.Show("Ошибка", "Не удалось обработать геометрию связных моделей");
                        return Result.Failed;
                    }

                    allSelectedElementsSolidsDiffusion = allSelectedElementsSolidsDiffusion
                        .Where(solid => solid != null && solid.Volume >= toleranceValue)
                        .ToList();

                    allSelectedLinkElementsSolidsDiffusion = allSelectedLinkElementsSolidsDiffusion
                        .Where(solid => solid != null && solid.Volume >= toleranceValue)
                        .ToList();

                    if (!allSelectedElementsSolidsDiffusion.Any() && !allSelectedLinkElementsSolidsDiffusion.Any())
                    {
                        TaskDialog.Show("Результат", "Не найдено подходящих SOLID.");
                        return Result.Failed;
                    }

                    using (Transaction tx = new Transaction(doc, "KPLN. Несовподения монолита"))
                    {
                        tx.Start();

                        FamilySymbol clashPointSymbol = GetClashPointSymbol(doc);
                        if (clashPointSymbol == null)
                        {
                            TaskDialog.Show("Ошибка", "Семейство 'ClashPoint' не найдено в проекте.");
                            tx.RollBack();
                            return Result.Failed;
                        }

                        foreach (Solid solid in allSelectedElementsSolidsDiffusion.Concat(allSelectedLinkElementsSolidsDiffusion))
                        {
                            if (solid == null || solid.Volume < 1e-5 || solid.Faces.IsEmpty)
                                continue;

                            XYZ center = solid.ComputeCentroid();
                            PlaceClashPoint(doc, center, clashPointSymbol);
                        }

                        tx.Commit();
                    }

                    IList<FamilyInstance> monolithClashPoints = GetMonolithClashPoints(doc);

                    if (monolithClashPoints.Count != 0)
                    {
                        List<(Element elem, string origin)> combinedSkipElementsTest =
                            skippedMainElements.Select(e => (e, "[Основной файл]. Возникла критическая проблема при создании общего Solid-элемента для дальнейшего логического вычитания."))
                                     .Concat(skippedLinkElements.Select(e => (e, "[Связный файл]. Возникла критическая проблема при создании общего Solid-элемента для дальнейшего логического вычитания.")))
                                     .ToList();

                        var win = new CheckMonolithInfo(uiDoc, monolithClashPoints, combinedSkipElementsTest)
                        {
                            Topmost = true,
                            ShowInTaskbar = true
                        };
                        win.Show();
                    }

                    return Result.Succeeded;
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    TaskDialog.Show("Информация", "Операция отменена пользователем");
                    return Result.Cancelled;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", $"{ex.Message}");
                    return Result.Failed;
                }
            }
            else
            {
                return Result.Failed;
            }
        }

        /// <summary>
        /// Метод для поиска геометрии
        /// </summary>
        private Solid GetUnionSolid(List<Element> elements, out List<Element> skippedElements)
        {
            Solid unionSolid = null;
            skippedElements = new List<Element>();

            foreach (Element el in elements)
            {
                bool elementSkipped = false;

                GeometryElement geomElement = el.get_Geometry(new Options() { ComputeReferences = true, IncludeNonVisibleObjects = true });
                if (geomElement == null) continue;

                foreach (GeometryObject geomObj in geomElement)
                {
                    IEnumerable<Solid> solids = ExtractSolidsFromGeometryObject(geomObj);
                    foreach (Solid solid in solids)
                    {
                        if (unionSolid == null)
                        {
                            unionSolid = solid;
                        }
                        else
                        {
                            try
                            {
                                unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);
                            }
                            catch
                            {
                                elementSkipped = true;
                            }
                        }
                    }
                }

                if (elementSkipped)
                    skippedElements.Add(el);
            }

            return unionSolid;
        }

        /// <summary>
        /// Метод для поиска геометрии (перегрузка)
        /// </summary>
        private Solid GetUnionSolid(List<(Element element, Transform transform)> elementsWithTransform, out List<Element> skippedElements)
        {
            Solid unionSolid = null;
            skippedElements = new List<Element>();

            foreach (var (el, transform) in elementsWithTransform)
            {
                bool elementSkipped = false;

                GeometryElement geomElement = el.get_Geometry(new Options() { ComputeReferences = true, IncludeNonVisibleObjects = true });
                if (geomElement == null) continue;

                foreach (GeometryObject geomObj in geomElement)
                {
                    IEnumerable<Solid> solids = ExtractSolidsFromGeometryObject(geomObj, transform);
                    foreach (Solid solid in solids)
                    {
                        if (unionSolid == null)
                        {
                            unionSolid = solid;
                        }
                        else
                        {
                            try
                            {
                                unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);
                            }
                            catch
                            {
                                elementSkipped = true;
                            }
                        }
                    }
                }

                if (elementSkipped)
                    skippedElements.Add(el);
            }

            return unionSolid;
        }

        /// <summary>
        /// Вспомогательный метод извлечения Solid
        /// </summary>
        private IEnumerable<Solid> ExtractSolidsFromGeometryObject(GeometryObject geomObj, Transform transform = null)
        {
            if (geomObj is Solid solid && solid.Volume > 0.1 && !solid.Faces.IsEmpty)
            {
                yield return transform != null ? SolidUtils.CreateTransformed(solid, transform) : solid;
            }
            else if (geomObj is GeometryInstance gi)
            {
                GeometryElement symbolGeometry = gi.GetSymbolGeometry();
                Transform instanceTransform = gi.Transform;
                Transform finalTransform = transform != null ? transform.Multiply(instanceTransform) : instanceTransform;

                foreach (GeometryObject obj in symbolGeometry)
                {
                    foreach (Solid nestedSolid in ExtractSolidsFromGeometryObject(obj, finalTransform))
                    {
                        yield return nestedSolid;
                    }
                }
            }
        }

        /// <summary>
        /// Получение ClashPoint
        /// </summary>
        public FamilySymbol GetClashPointSymbol(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfType<FamilySymbol>()
                .FirstOrDefault(f => f.Name == "ClashPoint");
        }

        /// <summary>
        /// Растановка ClashPoint
        /// </summary>
        public void PlaceClashPoint(Document doc, XYZ point, FamilySymbol clashPointType)
        {
            const double tol = 0.001;

            bool exists = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Any(fiOld =>
                {
                    if (fiOld.Symbol.Id != clashPointType.Id) return false;
                    LocationPoint lp = fiOld.Location as LocationPoint;
                    return lp != null && lp.Point.DistanceTo(point) < tol;
                });

            if (exists) return;

            if (!clashPointType.IsActive)
                clashPointType.Activate();

            FamilyInstance fi = doc.Create.NewFamilyInstance(
                 point, clashPointType, StructuralType.NonStructural);

            Parameter p = fi.LookupParameter("Комментарии");
            if (p != null && !p.IsReadOnly)
                p.Set("Проверка монолита");
        }

        /// <summary>
        /// Поиск созданных ClashPoint
        /// </summary>
        public static IList<FamilyInstance> GetMonolithClashPoints(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    if (fi.Symbol?.Family?.Name != "ClashPoint") return false;

                    string comments =
                        fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?
                          .AsString() ?? string.Empty;

                    return comments
                        .IndexOf("Проверка монолита", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();
        }
    }


    /// //////////////////////////////////////////////////////////
    // Внешний IExternalEventHandler. Обработчик удаления элемента
    internal class DeleteClashPointHandler : IExternalEventHandler
    {
        public ElementId ElementId { get; set; }

        public void Execute(UIApplication app)
        {
            if (ElementId == null || ElementId == ElementId.InvalidElementId) return;

            Document doc = app.ActiveUIDocument.Document;
            using (Transaction tx = new Transaction(doc, "KPLN. Удалить ClashPoint"))
            {
                tx.Start();
                doc.Delete(ElementId);
                tx.Commit();
            }
        }

        public string GetName() => nameof(DeleteClashPointHandler);
    }

    /// //////////////////////////////////////////////////////////
    // Внешний IExternalEventHandler. Создание обзорного 3D-вида
    internal class ShowClashPointHandler : IExternalEventHandler
    {
        public ElementId ElementId { get; set; }

        private static View3D RebuildPluginView(Document doc, View3D source3d)
        {
            const string PLUGIN_NAME = "3D. Проверка монолита (плагин)";

            View3D plugin = new FilteredElementCollector(doc)
                            .OfClass(typeof(View3D))
                            .Cast<View3D>()
                            .FirstOrDefault(v => !v.IsTemplate && v.Name == PLUGIN_NAME);

            if (plugin != null)
                return plugin;

            using (Transaction tx = new Transaction(doc, "KPLN. Создание служебного 3D-вида"))
            {
                tx.Start();
                ElementId dupId = source3d.Duplicate(ViewDuplicateOption.Duplicate);
                plugin = (View3D)doc.GetElement(dupId);
                plugin.Name = PLUGIN_NAME;
                tx.Commit();
            }
            return plugin;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (ElementId == null || ElementId == ElementId.InvalidElementId) return;

                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;

                View3D active3d = uidoc.ActiveView as View3D;
                if (active3d == null || active3d.IsTemplate) return;

                View3D v3 = RebuildPluginView(doc, active3d);

                using (Transaction tx = new Transaction(doc, "KPLN. Проверка монолита (ClahPoint на 3D-виде)"))
                {
                    tx.Start();

                    FamilyInstance fi = doc.GetElement(ElementId) as FamilyInstance;
                    if (fi == null) { tx.RollBack(); return; }

                    BoundingBoxXYZ bb = fi.get_BoundingBox(null);
                    if (bb == null) { tx.RollBack(); return; }



#if (Debug2020 || Revit2020)
                    double off = UnitUtils.ConvertToInternalUnits(2.0, DisplayUnitType.DUT_METERS);
#endif
#if (!Debug2020 && !Revit2020)
                    double off = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Meters);
#endif


                    v3.SetSectionBox(new BoundingBoxXYZ
                    {
                        Min = bb.Min - new XYZ(off, off, off),
                        Max = bb.Max + new XYZ(off, off, off)
                    });

                    XYZ center = fi.Location is LocationPoint lp
                                  ? lp.Point
                                  : (bb.Min + bb.Max) * 0.5;

                    XYZ forward = new XYZ(-0.25, -1, -0.15).Normalize();
                    XYZ right = forward.CrossProduct(XYZ.BasisZ).Normalize();
                    if (right.IsZeroLength())
                        right = forward.CrossProduct(XYZ.BasisX).Normalize();

                    XYZ up = right.CrossProduct(forward).Normalize();
                    XYZ eye = center - forward * off;

                    v3.SetOrientation(new ViewOrientation3D(eye, up, forward));

                    tx.Commit();
                }

                uidoc.ActiveView = v3;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ShowClashPointHandler",
                    $"Критическая ошибка:\n{ex.Message}");
            }
        }

        public string GetName() => nameof(ShowClashPointHandler);
    }

}