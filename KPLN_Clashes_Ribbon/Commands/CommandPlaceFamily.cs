using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    /// <summary>
    /// Класс для сравнения XYZ (для создания HashSet)
    /// </summary>
    internal sealed class ComparerByXYZ : IEqualityComparer<XYZ>
    {
        private const double _tolerance = 0.1;

        public bool Equals(XYZ x, XYZ y)
        {
            if (x == null || y == null)
                return false;

            return Math.Abs(x.DistanceTo(y)) < _tolerance;
        }

        public int GetHashCode(XYZ obj)
        {
            if (obj == null)
                return 0;

            // Заакругленне каардынат да хібнасці для стварэння стабільнага хэш-кода
            int hashX = (int)(obj.X / _tolerance);
            int hashY = (int)(obj.Y / _tolerance);
            int hashZ = (int)(obj.Z / _tolerance);

            // Камбінаванне хэш-кодаў каардынат
            return hashX ^ (hashY << 2) ^ (hashZ >> 2);
        }
    }

    public class CommandPlaceFamily : IExecutableCommand
    {
        public static readonly string FamilyName = "ClashPoint";

        private readonly ReportItem _report;

        public CommandPlaceFamily(ReportItem report)
        {
            _report = report;
        }

        public Result Execute(UIApplication app)
        {
            if (app.ActiveUIDocument == null)
                return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            try
            {
                // Проверка на 3д-вид
                if (!(uidoc.ActiveView is View3D activeView))
                {
                    TaskDialog.Show("Внимание", "Работать с метками можно только на 3D-виде. Открой 3D-вид");

                    return Result.Cancelled;
                }

                // Подготовка и создание клэшпоинтов
                using (Transaction t = new Transaction(doc, "KPLN_Указатель пересечения"))
                {
                    t.Start();


                    // Чистка от старых экз.
                    FamilyInstance[] oldFamInsOfGM = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_GenericModel)
                        .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == FamilyName)
                        .Cast<FamilyInstance>()
                        .ToArray();
                    ICollection<ElementId> availableWSOldElemsId = WorksharingUtils.CheckoutElements(doc, oldFamInsOfGM.Select(el => el.Id).ToArray());
                    doc.Delete(availableWSOldElemsId);


                    // Создание новых
                    FamilyInstance[] resultInst;
                    if (_report.SubElements.Any())
                    {
                        resultInst = GetClashPoints(uidoc, _report.SubElements.ToArray());
                        if (resultInst == null)
                            return Result.Cancelled;

                        // Выделяю элементы пересечения в модели
                        ElementId[] firstElemsId = _report.SubElements.Select(subRI => new ElementId(subRI.Element_1_Id)).ToArray();
                        ElementId[] secondElemsId = _report.SubElements.Select(subRI => new ElementId(subRI.Element_2_Id)).ToArray();
                        List<ElementId> elemsId = firstElemsId.Concat(secondElemsId).ToList();
                        uidoc.Selection.SetElementIds(elemsId);
                    }
                    else
                    {
                        resultInst = GetClashPoints(uidoc, new ReportItem[1] { _report });
                        if (resultInst == null)
                            return Result.Cancelled;

                        // Выделяю элементы пересечения в модели
                        uidoc.Selection.SetElementIds(new List<ElementId>
                        {
                            new ElementId(_report.Element_1_Id),
                            new ElementId(_report.Element_2_Id)
                        });
                    }


                    // Приближаю к элементу
                    ZoomTools.ZoomElement(SumBBox(resultInst), app.ActiveUIDocument, activeView);

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Cancelled;
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

        /// <summary>
        /// Создание коллекции FamilyInstance точек пересечения
        /// </summary>
        private static FamilyInstance[] GetClashPoints(UIDocument uidoc, ReportItem[] riArr)
        {
            Document doc = uidoc.Document;
            HashSet<XYZ> xyzToCreate = new HashSet<XYZ>(new ComparerByXYZ());

            Transform docTRans = doc.ActiveProjectLocation.GetTotalTransform();
            
            //// Уточняю трансофрм, если БТП была смещена
            //var docBP = new FilteredElementCollector(doc)
            //    .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
            //    .FirstOrDefault();
            //var docBPBBox = docBP.get_BoundingBox(null);
            ////if (!docBPBBox.Max.IsAlmostEqualTo(XYZ.Zero, 0.1))
            ////    docTRans *= (Transform.CreateTranslation(docBPBBox.Max).Inverse);

            // Создание новых
            // Метка, что хотя бы один элемент был удален
            bool elemsDeleted = false;
            bool elemsNotInOpenFile = true;
            foreach (ReportItem subRI in riArr)
            {
                // Проверка на наличие элемента в файле
                Element elem1 = doc.GetElement(new ElementId(subRI.Element_1_Id));
                Element elem2 = doc.GetElement(new ElementId(subRI.Element_2_Id));
                if (elem1 == null && elem2 == null)
                {
                    if (doc.Title.Contains(subRI.Element_1_DocName) || doc.Title.Contains(subRI.Element_2_DocName))
                        elemsDeleted = true;
                }
                else
                    elemsNotInOpenFile = false;

                XYZ subRIPoint = GetXYZFromReportItem(subRI);
                XYZ transLoc = docTRans.OfPoint(subRIPoint);

                xyzToCreate.Add(transLoc);
            }

            if (elemsNotInOpenFile)
            {
                TaskDialog.Show("Внимание!",
                    "Элементы отсутсвуют в открытом проекте - скорее всего это элементы не относятся к вашей модели.\n" +
                    "ВАЖНО: Внимательно прочитай имя отчета - оно должно содержать аббревиатуру твоего раздела");

                return null;
            }

            if (elemsDeleted)
                TaskDialog.Show("Внимание!",
                    "Элемент или группа элементов - отсутсвуют в открытом проекте, т.к. скорее всего они были удалены.\n" +
                    "ВАЖНО: Точка пересечения всё равно появится, чтобы избежать пропуска перемоделированных элементов");

            FamilyInstance[] result = new FamilyInstance[xyzToCreate.Count];
            for (int i = 0; i < xyzToCreate.Count; i++)
            {
                result[i] = CreateFamilyInstance(doc, xyzToCreate.ElementAt(i));
            }

            return result.ToArray();
        }

        /// <summary>
        /// Создать общий BoundingBoxXYZ для элементов
        /// </summary>
        private static BoundingBoxXYZ SumBBox(IEnumerable<Element> elems)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (Element element in elems)
            {
                BoundingBoxXYZ elementBox = element.get_BoundingBox(null);
                if (elementBox == null)
                    continue;

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = elementBox.Min,
                        Max = elementBox.Max
                    };
                }
                else
                {
                    resultBBox.Min = new XYZ(
                        Math.Min(resultBBox.Min.X, elementBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elementBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elementBox.Min.Z));

                    resultBBox.Max = new XYZ(
                        Math.Max(resultBBox.Max.X, elementBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elementBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elementBox.Max.Z));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Получить точку XYZ из ReportItem
        /// </summary>
        private static XYZ GetXYZFromReportItem(ReportItem ri)
        {
            {
                string pt = ri.Point;
                pt = pt.Replace("X:", "");
                pt = pt.Replace("Y:", "");
                pt = pt.Replace("Z:", "");

                string pts = string.Empty;
                foreach (char c in pt)
                {
                    if ("-0123456789.,".Contains(c))
                    {
                        pts += c;
                    }
                }

                string[] parts = pts.Split(',');
                if (
                    double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointX)
                    && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointY)
                    && double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointZ))
                    return new XYZ(pointX * 3.28084, pointY * 3.28084, pointZ * 3.28084);
                else
                    throw new Exception("Проблемы с CultureInfo");
            }
        }

        private static FamilyInstance CreateFamilyInstance(Document doc, XYZ point)
        {
            Level level = GetNearestLevel(doc, point.Z) ?? throw new Exception("В проекте отсутсвуют уровни!");

            FamilySymbol intersectFamSymb = GetFamilySymbol(doc);

            FamilyInstance instance = doc
                .Create
                .NewFamilyInstance(point, intersectFamSymb, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(point.Z - level.Elevation);
            doc.Regenerate();

            return instance;
        }

        private static FamilySymbol GetFamilySymbol(Document doc)
        {
            FamilySymbol[] oldFamSymbOfGM = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == FamilyName)
                .Cast<FamilySymbol>()
                .ToArray();

            // Если в проекте нет - то грузим
            if (!oldFamSymbOfGM.Any())
            {
                string path = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{ModuleData.RevitVersion}\{FamilyName}.rfa";
                bool result = doc.LoadFamily(path);
                if (!result)
                    throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                doc.Regenerate();

                // Повторяем после загрузки семейства
                oldFamSymbOfGM = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == FamilyName)
                    .Cast<FamilySymbol>()
                    .ToArray();
            }

            FamilySymbol searchSymbol = oldFamSymbOfGM.FirstOrDefault();
            searchSymbol.Activate();

            return searchSymbol;
        }
    }
}
