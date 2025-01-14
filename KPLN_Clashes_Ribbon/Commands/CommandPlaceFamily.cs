using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandPlaceFamily : IExecutableCommand
    {
        private static readonly string _familyName = "ClashPoint";
        private static FamilySymbol _intersectFamSymb;

        private readonly ReportWindow _reportWindow;

        private readonly XYZ _point;
        private readonly int _id1;
        private readonly string _elementInfo1;
        private readonly int _id2;
        private readonly string _elementInfo2;

        public CommandPlaceFamily(XYZ point, int id1, string info1, int id2, string info2, ReportWindow window)
        {
            _reportWindow = window;
            _id1 = id1;
            _elementInfo1 = info1;
            _id2 = id2;
            _elementInfo2 = info2;
            _point = new XYZ(point.X * 3.28084, point.Y * 3.28084, point.Z * 3.28084);
        }
        
        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            try
            {
                if (app.ActiveUIDocument != null)
                {
                    // Проверка на наличие элемента в файле
                    Element element1 = doc.GetElement(new ElementId(_id1));
                    Element element2 = doc.GetElement(new ElementId(_id2));
                    if (element1 == null && element2 == null)
                    {
                        TaskDialog.Show("Внимание!", "Все элементы в связи. Метку необходимо расставлять в документах, где элемент присутсвует (см. информацию об элементах отчета)");
                        return Result.Cancelled;
                    }
                    
                    // Проверка на совпадение элемента в файле и элемента в отчете
                    bool el1 = true;
                    bool el2 = true;
                    if (element1 != null)
                    {
                        if (!_elementInfo1.Contains(element1.Name) && !_elementInfo2.Contains(element1.Name))
                        {
                            el1 = false;
                        }
                    }
                    if (element2 != null)
                    {
                        if (!_elementInfo1.Contains(element2.Name) && !_elementInfo2.Contains(element2.Name))
                        {
                            el2 = false;
                        }
                    }
                    if (!el1 && !el2)
                    {
                        TaskDialog.Show("Внимание!", "Имя данного элемента в отчете не совпадает с именем элемента в проекте. " +
                            "Произошла подмена id-элементов, коллизию нужно уточнить у проверяющего");
                        return Result.Cancelled;
                    }
                    
                    using (Transaction t = new Transaction(doc, "Указатель"))
                    {
                        t.Start();

                        // Чистка от старых экз.
                        FamilyInstance[] oldFamInsOfGM = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_GenericModel)
                            .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == _familyName)
                            .Cast<FamilyInstance>()
                            .ToArray();
                        ICollection<ElementId> availableWSOldElemsId = WorksharingUtils.CheckoutElements(doc, oldFamInsOfGM.Select(el => el.Id).ToArray());
                        doc.Delete(availableWSOldElemsId);

                        // Создание новых
                        XYZ transformed_location = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(_point);
                        FamilyInstance createdinstance = CreateFamilyInstance(doc, transformed_location);
                    
                        // Выделяю элементы пересечения в модели
                        if (element1 != null && element2 != null)
                            uidoc.Selection.SetElementIds(new List<ElementId>() { element1.Id, element2.Id });
                        else if (element1 != null)
                            uidoc.Selection.SetElementIds(new List<ElementId>() { element1.Id });
                        else if (element2 != null)
                            uidoc.Selection.SetElementIds(new List<ElementId>() { element2.Id });
                        // Не должно происходить, но для стабильности добавил
                        else
                            uidoc.Selection.SetElementIds(new List<ElementId>() { createdinstance.Id });

                        if (createdinstance != null) 
                            _reportWindow.OnClosingActions.Add(new CommandRemoveInstance(doc, createdinstance));
                    
                        ZoomTools.ZoomElement(createdinstance.get_BoundingBox(null), app.ActiveUIDocument);
                    
                        t.Commit();
                    }
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }
        /// <summary>
        /// Получить коллекцию ВСЕХ уровней проекта
        /// </summary>
        public static Level[] GetLevels(Document doc)
        {
            Level[] instances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToArray();

            return instances;
        }

        /// <summary>
        ///  Поиск ближайшего подходящего уровня
        /// </summary>
        private static Level GetNearestLevel(Document doc, double elevation)
        {
            Level result = null;

            double resultDistance = 999999;
            foreach (Level lvl in GetLevels(doc))
            {
                double tempDistance = Math.Abs(lvl.Elevation - elevation);
                if (Math.Abs(lvl.Elevation - elevation) < resultDistance)
                {
                    result = lvl;
                    resultDistance = tempDistance;
                }
            }
            return result;
        }

        private static FamilyInstance CreateFamilyInstance(Document doc, XYZ point)
        {
            Level level = GetNearestLevel(doc, point.Z) ?? throw new Exception("В проекте отсутсвуют уровни!");

            FamilySymbol intersectFamSymb = GetFamilySymbol(doc);

            // Чистка от старых экз.
            FamilyInstance[] oldFamInsOfGM = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == _familyName)
                .Cast<FamilyInstance>()
                .ToArray();

            FamilyInstance oldEqualFI = oldFamInsOfGM
                .FirstOrDefault(old => 
                    old.Location is LocationPoint oldLocPnt 
                    && Math.Abs(oldLocPnt.Point.DistanceTo(point)) < 0.05);
            if (oldEqualFI != null)
                return oldEqualFI;

            FamilyInstance instance = doc
                .Create
                .NewFamilyInstance(point, intersectFamSymb, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            
            doc.Regenerate();
            instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(point.Z - level.Elevation);
            doc.Regenerate();
            
            return instance;
        }
        
        private static FamilySymbol GetFamilySymbol(Document doc)
        {
            FamilySymbol[] oldFamSymbOfGM = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == _familyName)
                .Cast<FamilySymbol>()
                .ToArray();

            // Если в проекте нет - то грузим
            if (!oldFamSymbOfGM.Any())
            {
                string path = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{ModuleData.RevitVersion}\{_familyName}.rfa";
                bool result = doc.LoadFamily(path);
                if (!result)
                    throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                doc.Regenerate();

                // Повторяем после загрузки семейства
                oldFamSymbOfGM = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == _familyName)
                    .Cast<FamilySymbol>()
                    .ToArray();
            }

            FamilySymbol searchSymbol = oldFamSymbOfGM.FirstOrDefault();
            searchSymbol.Activate();

            return searchSymbol;
        }
    }
}
