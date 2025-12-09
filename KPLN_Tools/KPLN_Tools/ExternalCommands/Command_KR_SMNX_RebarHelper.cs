using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    internal sealed class KGConstruction 
    {
        private Solid[] _constrSolid;

        public KGConstruction(Element elem, string code, double volume)
        {
            ConstrElem = elem;
            ConstrCode = code;
            ConstrVolume = volume;
        }

        internal Element ConstrElem { get; }

        internal string ConstrCode { get; }

        internal double ConstrVolume { get; }

        internal Solid[] ConstrSolidColl 
        {
            get
            {
                if (_constrSolid == null)
                    _constrSolid = GeomUtil.GetElementSolids(ConstrElem);

                return _constrSolid;
            }
        }
    }

    /// <summary>
    /// Работа с геомтерией
    /// </summary>
    internal static class GeomUtil
    {
        internal static Solid[] GetElementSolids(Element element)
        {
            List<Solid> solidColl = new List<Solid>();

            Options opt = new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = true
            };
            GeometryElement geomElem = element.get_Geometry(opt);
            if (geomElem != null)
            {
                GetSolidsFromGeomElem(geomElem, Transform.Identity, solidColl);
            }

            BoundingBoxXYZ elemBbox = element.get_BoundingBox(null);
            Solid[] solidWithVolume = solidColl.Where(s => s.Volume > 0).ToArray();
            if (solidWithVolume.Length > 5 && elemBbox != null)
                return new Solid[1] { CreateSolidFromBoundingBox(elemBbox) };

            return solidWithVolume;
        }

        /// <summary>
        /// Получить Solid из элемента
        /// </summary>
        private static void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solids)
        {
            foreach (GeometryObject geomObject in geometryElement)
            {
                switch (geomObject)
                {
                    case Solid solid:
                        solids.Add(solid);
                        break;

                    case GeometryInstance geomInstance:
                        GetSolidsFromGeomElem(geomInstance.GetInstanceGeometry(), geomInstance.Transform.Multiply(transformation), solids);
                        break;

                    case GeometryElement geomElem:
                        GetSolidsFromGeomElem(geomElem, transformation, solids);
                        break;
                }
            }
        }

        private static Solid CreateSolidFromBoundingBox(BoundingBoxXYZ bbox)
        {
            // Мінімальныя і максімальныя каардынаты ў bbox
            XYZ min = bbox.Min;
            XYZ max = bbox.Max;

            // Улічваем трансфармацыю BoundingBox
            Transform t = bbox.Transform;

            // Ствараем 8 кропак паралелепіпеда
            XYZ p0 = t.OfPoint(new XYZ(min.X, min.Y, min.Z));
            XYZ p1 = t.OfPoint(new XYZ(max.X, min.Y, min.Z));
            XYZ p2 = t.OfPoint(new XYZ(max.X, max.Y, min.Z));
            XYZ p3 = t.OfPoint(new XYZ(min.X, max.Y, min.Z));

            XYZ p4 = t.OfPoint(new XYZ(min.X, min.Y, max.Z));
            XYZ p5 = t.OfPoint(new XYZ(max.X, min.Y, max.Z));
            XYZ p6 = t.OfPoint(new XYZ(max.X, max.Y, max.Z));
            XYZ p7 = t.OfPoint(new XYZ(min.X, max.Y, max.Z));

            // Ніжняя грань
            List<Curve> baseLoop = new List<Curve>()
            {
                Line.CreateBound(p0, p1),
                Line.CreateBound(p1, p2),
                Line.CreateBound(p2, p3),
                Line.CreateBound(p3, p0)
            };

            CurveLoop curveLoop = CurveLoop.Create(baseLoop);

            // Напрамак экструзіі (па Z)
            XYZ extrusionDir = (p4 - p0).Normalize();
            double height = (p4 - p0).GetLength();

            // Ствараем solid праз экструзію
            Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { curveLoop },
                extrusionDir,
                height);

            return solid;
        }
    }


    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_KR_SMNX_RebarHelper : IExternalCommand
    {
        private readonly static string _constrZhBMark = "Марка";
        private readonly static string _constrRebarMark = "Мрк.МаркаКонструкции";
        private readonly static string _constrRebarHostMark = "Метка основы";

        /// <summary>
        /// Это исключения из категорий ЖБ элементов. Нужен для подсчета объема бетона. 
        /// Список исключений в именах СЕМЕЙСТВ И ТИПОВ ревит (НАЧИНАЕТСЯ С) для генерации исключений в выбранных категориях.
        /// ВАЖНО: Имя семейства и типа в ревит прописано в параметре ELEM_FAMILY_AND_TYPE_PARAM, и формируется в формате "Имя семейства" + ": " + "Имя типа"
        /// </summary>
        private readonly List<string> _exceptFamAndTypeNameStartWithList_ZhB = new List<string>
        {
            "Базовая стена: 01_",
            "220_",
            "221_",
            "225_",
            "232_",
            "233_",
            "263_",
            "264_",
            "265_",
            "352_",
            "353_",
            "360_",
            "370_",
            "Обозначение участка с отгибом",
        };

        /// <summary>
        /// Это исключения из арматурных элементов.
        /// Список исключений в именах СЕМЕЙСТВ И ТИПОВ ревит (НАЧИНАЕТСЯ С) для генерации исключений в выбранных категориях.
        /// ВАЖНО: Имя семейства и типа в ревит прописано в параметре ELEM_FAMILY_AND_TYPE_PARAM, и формируется в формате "Имя семейства" + ": " + "Имя типа"
        /// </summary>
        private readonly List<string> _exceptFamAndTypeNameStartWithList_Rb = new List<string>
        {
            "220_",
            "352_",
            "353_",
        };

        private readonly List<KGConstruction> _kgConstrColl= new List<KGConstruction>();
        private readonly Dictionary<ElementId, List<Element>> _hostForRebarsDict = new Dictionary<ElementId, List<Element>>();
        private readonly Dictionary<string, List<Element>> _fatalErrorConstrElemsDict = new Dictionary<string, List<Element>>();
        private readonly Dictionary<string, List<Element>> _errorConstrElemsDict = new Dictionary<string, List<Element>>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region Генерация фильтров
            // Фильтрый ЖБ по имени
            List<FilterRule> filtRules_ZhB = new List<FilterRule>();
            foreach (string currentName in _exceptFamAndTypeNameStartWithList_ZhB)
            {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM), currentName, true);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM), currentName);
#endif
                filtRules_ZhB.Add(fRule);
            }
            ElementParameterFilter eFilter_ZhB = new ElementParameterFilter(filtRules_ZhB);

            List<ElementFilter> reultFilters = new List<ElementFilter>();
            
            // Фильтры Арматуры по имени
            List<FilterRule> filtRules_Rb = new List<FilterRule>();
            foreach (string currentName in _exceptFamAndTypeNameStartWithList_Rb)
            {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM), currentName, true);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM), currentName);
#endif
                filtRules_Rb.Add(fRule);
            }
            ElementParameterFilter eFilter_Rb = new ElementParameterFilter(filtRules_Rb);
            reultFilters.Add(eFilter_Rb);

            // Фильтры Арматуры по РН
            List<Workset> worksets = new FilteredWorksetCollector(doc).ToList();
            
            Workset raspSyst = worksets.Where(w => w.Name.Equals("КР_Распорная система")).FirstOrDefault();
            if (raspSyst != null)
                reultFilters.Add(new ElementWorksetFilter(raspSyst.Id, true));
            
            Workset shp = worksets.Where(w => w.Name.Equals("КР_Шпунт")).FirstOrDefault();
            if (shp != null)
                reultFilters.Add(new ElementWorksetFilter(shp.Id, true));


            // Результирующий фильтр для Арматуры
            LogicalAndFilter wsResultFilter = new LogicalAndFilter(reultFilters);
            #endregion

            TaskDialog taskDialog = new TaskDialog("Выбери действие")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainContent = "Записать объём бетона и основную марку в арматуру - нажми Да.\nПеренести значения из txt в параметры бетонных конструкций - нажми Повтор.",
                CommonButtons = TaskDialogCommonButtons.Retry | TaskDialogCommonButtons.Yes
            };
            
            TaskDialogResult userInput = taskDialog.Show();
            if (userInput == TaskDialogResult.Cancel)
                return Result.Cancelled;
            else if (userInput == TaskDialogResult.Yes)
            {
                List<BuiltInCategory> constrCats = new List<BuiltInCategory> {
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_StructuralFoundation
                };

                List<Element> constrs = new FilteredElementCollector(doc)
                    .WherePasses(new ElementMulticategoryFilter(constrCats))
                    .WherePasses(eFilter_ZhB)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();
                foreach (Element constr in constrs)
                {
                    Parameter markParam = constr.LookupParameter(_constrZhBMark);
                    if (markParam == null || !markParam.HasValue)
                    {
                        string errorMsg = "Марка у бетона не заполнена";
                        if (_errorConstrElemsDict.ContainsKey(errorMsg))
                            _errorConstrElemsDict[errorMsg].Add(constr);
                        else
                            _errorConstrElemsDict.Add(errorMsg, new List<Element>() { constr });
                    }
                    else
                    {
                        string mark = markParam.AsString();

                        Parameter volumeParam = constr.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                        if (volumeParam == null || !volumeParam.HasValue)
                        {
                            string errorMsg = "Элемент не имеет объёма";
                            if (_errorConstrElemsDict.ContainsKey(errorMsg))
                                _errorConstrElemsDict[errorMsg].Add(constr);
                            else
                                _errorConstrElemsDict.Add(errorMsg, new List<Element>() { constr });
                        }
#if Debug2020 || Revit2020
                        double volume = Math.Round(UnitUtils.ConvertFromInternalUnits(volumeParam.AsDouble(), DisplayUnitType.DUT_CUBIC_METERS), 3);
#else
                        double volume = Math.Round(UnitUtils.ConvertFromInternalUnits(
                            volumeParam.AsDouble(), 
                            new ForgeTypeId("autodesk.revit.parameter:hostVolumeComputed-1.0.0")), 3);
#endif

                        string constrCode;
                        if (constr is Wall wall)
                        {
                            string subTypeMark = wall.LookupParameter("Орг.ПодтипЭлемента").AsString();
                            constrCode = $"{mark + subTypeMark}~{volume}~{constr.Id}";
                        }
                        else
                            constrCode = $"{mark}~{volume}~{constr.Id}";

                        _kgConstrColl.Add(new KGConstruction(constr, constrCode, volume));
                    }
                }

                List<Element> rebars = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rebar)
                    .WherePasses(wsResultFilter)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                foreach (Element rebar in rebars)
                {
                    int index = rebars.IndexOf(rebar);
                    PrepareHostForRebars(doc, rebar);
                }

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("SMNX: Зпись данных в арматуру");

                    SetParamData(doc);

                    if (_fatalErrorConstrElemsDict.Count > 0)
                    {
                        foreach (KeyValuePair<string, List<Element>> kvp in _errorConstrElemsDict)
                        {
                            string elemIdColl = string.Join(",", kvp.Value.Select(e => e.Id.ToString()));
                            HtmlOutput.Print(
                                $"Критическая ошибка (работа остановлена): \"{kvp.Key}\" для элементов:\n {elemIdColl}",
                                MessageType.Warning);
                        }

                        t.RollBack();
                    }
                    else
                    {
                        if (_errorConstrElemsDict.Count > 0)
                        {
                            foreach (KeyValuePair<string, List<Element>> kvp in _errorConstrElemsDict)
                            {
                                string elemIdColl = string.Join(",", kvp.Value.Select(e => e.Id.ToString()));
                                HtmlOutput.Print(
                                    $"Ошибка: \"{kvp.Key}\" для элементов:\n {elemIdColl}",
                                    MessageType.Warning);
                            }
                        }

                        t.Commit();
                    }
                }
            }
            else
            {
                DataParseToRevit(doc);
            }

            return Result.Succeeded;
        }

        private void PrepareHostForRebars(Document doc, Element rebarElement)
        {
            Parameter isIfcRebarParam = rebarElement.LookupParameter("Арм.ВыполненаСемейством");
            if (isIfcRebarParam == null)
            {
                ElementId typeId = rebarElement.GetTypeId();
                if (typeId == ElementId.InvalidElementId) return;

                Element typeElem = doc.GetElement(typeId);
                if (typeElem == null) return;
                isIfcRebarParam = typeElem.LookupParameter("Арм.ВыполненаСемейством");

                if (isIfcRebarParam == null) return;
            }

            int isIfc = isIfcRebarParam.AsInteger();
            Parameter hostNameParam;
            Element hostElement = null;
            if (isIfc == 1)
                hostNameParam = rebarElement.LookupParameter(_constrRebarMark);
            else
            {
                hostNameParam = rebarElement.LookupParameter(_constrRebarHostMark);
                if (rebarElement is Rebar rebar)
                    hostElement = doc.GetElement(rebar.GetHostId());
                else if (rebarElement is RebarInSystem rebarInSystem)
                    hostElement = doc.GetElement(rebarInSystem.GetHostId());
            }

            if (hostNameParam == null || !hostNameParam.HasValue)
            {
                string errorMsg = "Основа не определена";
                if (_errorConstrElemsDict.ContainsKey(errorMsg))
                    _errorConstrElemsDict[errorMsg].Add(rebarElement);
                else
                    _errorConstrElemsDict.Add(errorMsg, new List<Element>() { rebarElement });
            }
            else
            {
                // Попытка поиска основания по связям ревит
                bool isErorMark = true;
                if (hostElement != null)
                    isErorMark = false;
                else if (rebarElement is FamilyInstance rebarFI)
                {
                    Element superComp = rebarFI.SuperComponent;
                    if (superComp != null && superComp is FamilyInstance superCompFI)
                    {
                        Element superCompHost = superCompFI.Host;
                        if (superCompHost != null)
                        {
                            isErorMark = false;
                            hostElement = superCompHost;
                        }
                    }
                }

                string hostMark = hostNameParam.AsString();
                if (string.IsNullOrEmpty(hostMark)) 
                    return;

                // Обнуляю хост если это уровень
                if (hostElement != null && hostElement is Level _)
                    hostElement = null;
                
                double tempIntersectSolidVolume = 0;
                // Попытка поиска основания по геометрии
                if (hostElement == null)
                {
                    Solid[] rebarSolids = GeomUtil.GetElementSolids(rebarElement);
                    // Игнор родительских пустых семейств
                    if (!rebarSolids.Any()) return;

                    foreach (KGConstruction kgConstr in _kgConstrColl)
                    {
                        if (!kgConstr.ConstrCode.Split('~')[0].StartsWith(hostMark)) continue;
                        
                        isErorMark = false;
                         
                        // Игнор пустых основ
                        if (!kgConstr.ConstrSolidColl.Any()) return;

                        try
                        {
                            // Попытка поиска основания по Solid (точное положение)
                            foreach (Solid rebarSolid in rebarSolids)
                            {
                                foreach (Solid elemSolid in kgConstr.ConstrSolidColl)
                                {
                                    Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(rebarSolid, elemSolid, BooleanOperationsType.Intersect);
                                    if (intersectionSolid != null && intersectionSolid.Volume > tempIntersectSolidVolume)
                                    {
                                        tempIntersectSolidVolume = intersectionSolid.Volume;
                                        hostElement = kgConstr.ConstrElem;
                                    }
                                }
                            }

                            // Попытка поиска основания по притягиванию по Z (пониженная точность)
                            if (hostElement == null)
                            {
                                foreach (Solid rebarSolid in rebarSolids)
                                {
                                    Transform rebarTransform = rebarSolid.GetBoundingBox().Transform;
                                    Transform rebarInverseTransform = rebarTransform.Inverse;
                                    Solid rebarZerotransformSolid = SolidUtils.CreateTransformed(rebarSolid, rebarInverseTransform);
                                    foreach (Solid elemSolid in kgConstr.ConstrSolidColl)
                                    {
                                        XYZ elemCentroid = elemSolid.ComputeCentroid();
                                        rebarTransform.Origin = new XYZ(rebarTransform.Origin.X, rebarTransform.Origin.Y, elemCentroid.Z);

                                        Solid transformedByElemRebarSolid = SolidUtils.CreateTransformed(rebarZerotransformSolid, rebarTransform);
                                        Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(elemSolid, transformedByElemRebarSolid, BooleanOperationsType.Intersect);
                                        if (intersectionSolid != null && intersectionSolid.Volume > tempIntersectSolidVolume)
                                        {
                                            tempIntersectSolidVolume = intersectionSolid.Volume;
                                            hostElement = kgConstr.ConstrElem;
                                        }
                                    }
                                }
                            }
                        }
                        //Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            continue;
                        }
                    }
                }

                if (isErorMark)
                {
                    // Обход проблем с оформлением в подземке - мы их не учитываем, а в стандарте - блочим проверку
                    if (doc.PathName.Contains("_КР_КЖ_UN_"))
                    {
                        string errorMsg = "Марка арматуры не совпадает ни с одной макрой ЖБ элемента";
                        if (_errorConstrElemsDict.ContainsKey(errorMsg))
                            _errorConstrElemsDict[errorMsg].Add(rebarElement);
                        else
                            _errorConstrElemsDict.Add(errorMsg, new List<Element>() { rebarElement });
                    }
                    else
                    {
                        string errorMsg = "Марка арматуры не совпадает ни с одной макрой ЖБ элемента";
                        if (_fatalErrorConstrElemsDict.ContainsKey(errorMsg))
                            _fatalErrorConstrElemsDict[errorMsg].Add(rebarElement);
                        else
                            _fatalErrorConstrElemsDict.Add(errorMsg, new List<Element>() { rebarElement });
                    }
                }

                if (hostElement != null)
                {
                    if (_hostForRebarsDict.ContainsKey(hostElement.Id))
                        _hostForRebarsDict[hostElement.Id].Add(rebarElement);
                    else
                        _hostForRebarsDict.Add(hostElement.Id, new List<Element>() { rebarElement });
                }
                else
                {
                    string errorMsg = "Не удалось определить основу для арматуры";
                    if (_errorConstrElemsDict.ContainsKey(errorMsg))
                        _errorConstrElemsDict[errorMsg].Add(rebarElement);
                    else
                        _errorConstrElemsDict.Add(errorMsg, new List<Element>() { rebarElement });
                }
            }
        }

        private void SetParamData(Document doc)
        {
            foreach (KeyValuePair<ElementId, List<Element>> kvp in _hostForRebarsDict)
            {
                Element hostElement = doc.GetElement(kvp.Key);
                foreach (Element rebarElement in kvp.Value)
                {
                    KGConstruction kgConstr = _kgConstrColl.Where(constr => constr.ConstrCode.EndsWith(hostElement.Id.ToString())).FirstOrDefault();
                    if (kgConstr == null)
                        continue;

                    rebarElement.LookupParameter("СоставнаяМарка").Set(kgConstr.ConstrCode);

                    double volumeCubicM = kgConstr.ConstrVolume;
#if Debug2020 || Revit2020
                    double volume = Math.Round(UnitUtils.ConvertToInternalUnits(volumeCubicM, DisplayUnitType.DUT_CUBIC_METERS), 3);
#else
                    double volume = Math.Round(UnitUtils.ConvertToInternalUnits(
                        volumeCubicM,
                        new ForgeTypeId("autodesk.revit.parameter:hostVolumeComputed-1.0.0")), 3);
#endif

                    string rebarVolumeParamName = "Арм.ОбъемБетона";
                    Parameter volumeParam = rebarElement.LookupParameter(rebarVolumeParamName);
                    if (volumeParam == null)
                    {
                        TaskDialog.Show("Ошибка", "Нет параметра " + rebarVolumeParamName);
                        throw new Exception();
                    }
                    volumeParam.Set(volume);
                }
            }
        }

        private void DataParseToRevit(Document doc)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"Y:\Жилые здания\Обыденский\10.Стадия_Р\6.КР\1.RVT\1.Расчет металлоёмкости", // Начальная директория
                Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*" // Фильтр файлов
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("SMNX: Зпись данных в монолит");
                    try
                    {
                        string filePath = openFileDialog.FileName;

                        // Чтение всех строк из файла
                        string[] lines = File.ReadAllLines(filePath);

                        // Обработка каждой строки
                        for (int i = 2; i < lines.Count(); i++)
                        {
                            string line = lines[i];
                            string hostIdData = line.Split('	')[3].Split('~')[2];
                            string hostVolumeData = line.Split('	')[2];
                            if (string.IsNullOrEmpty(hostIdData) || string.IsNullOrEmpty(hostVolumeData))
                                throw new Exception($"В строке {line} не удалось получить данные. Покажи разработчику");

                            Element hostElem = null;
                            if (int.TryParse(hostIdData, out int elemId))
                            {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                                int hostId = int.Parse(hostIdData);
#else
                                long hostId = long.Parse(hostIdData);
#endif
                                hostElem = doc.GetElement(new ElementId(hostId));
                                if (hostElem == null)
                                    throw new Exception($"Элемент с id: {hostIdData} отсутсвует. Покажи разработчику");
                            }
                            else
                                throw new Exception($"Распарсить id: {hostIdData} не удалось. Покажи разработчику");

                            if (double.TryParse(hostVolumeData, out double volume))
                            {
                                hostElem.get_Parameter(new Guid("c12156d8-db24-4719-93cc-4c87ac359906")).Set(volume);
                            }
                            else
                                throw new Exception($"Не удалось преобразовать данные в double: {hostVolumeData}. Покажи разработчику");
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Ошибка", $"Покажи разработчику: {ex.Message}");
                        throw new Exception();
                    }
                    t.Commit();
                }
            }
        }
    }
}