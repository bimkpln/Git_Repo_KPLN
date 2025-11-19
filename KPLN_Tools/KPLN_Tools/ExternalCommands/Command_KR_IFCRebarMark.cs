using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace KPLN_Tools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_KR_IFCRebarMark : IExternalCommand
    {
        internal const string PluginName = "Маркировка IFC арматуры";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;   
            Document doc = uidoc.Document;

            FilteredElementCollector RebarIFCFilter = new FilteredElementCollector(doc, doc.ActiveView.Id); //формируем фильтр

            List<ElementId> parentIds = new List<ElementId>();    //список ID всех размещенных родительских семейств IFC или просто стержней IFC
            List<Element> IFCRebars = new List<Element>();  //список размещенных стержней IFC арматуры, похож на parentIds

            List<Element> TempIFCRebars = new List<Element>();  //временное хранение стержней (для фильтрации)
            List<ElementId> IFCRebarIDs = new List<ElementId>();  //список ID размещенных стержней IFC арматуры 

            Dictionary<Element, List<Element>> rebarToStruct = new Dictionary<Element, List<Element>>(); //словарь по типу арматура-список возможных монолитных хостов


            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();

                #region Составление списка IFC-арматур (отдельные стержни + родительские с-ва проемов, отверстий и т.п.)
                IFCRebars = RebarIFCFilter.OfCategory(BuiltInCategory.OST_Rebar)
                                              .WhereElementIsNotElementType()
                                              .ToElements()
                                              .ToList();

                //среди них оставляем только IFC арматуру
                foreach (Element rebar in IFCRebars)
                {
                    ElementType rebarType = doc.GetElement(rebar.GetTypeId()) as ElementType;
                    Parameter paramIsIFC = rebarType?.LookupParameter("Арм.ВыполненаСемейством");
                    if (paramIsIFC != null) //Проверка, что параметр найден
                    {
                        int IsIFC = paramIsIFC.AsInteger();
                        if (IsIFC == 1)
                        {
                            TempIFCRebars.Add(rebar);
                        }
                    }
                }
                IFCRebars = TempIFCRebars;

                IFCRebarIDs = IFCRebars.Select(it => it.Id).ToList();   //список ID стержней IFC арматуры

                foreach (ElementId elid in IFCRebarIDs)
                {
                    Element el = doc.GetElement(elid);
                    ElementId parentId = null;


                    if (el is FamilyInstance)     //для ifc арматуры и семейств по типу окно, дверь и т.п. 
                    {
                        FamilyInstance elfam = el as FamilyInstance;
                        Element parentFamily = elfam.SuperComponent;
                        if (parentFamily != null)       //если стержень это часть вложенного с-ва
                            parentId = parentFamily.Id;
                        else                            //иначе возвращаем просто стержень
                            parentId = elfam.Id;
                    }
                    if (parentId != null)
                    {
                        if (!parentIds.Contains(parentId))
                            parentIds.Add(parentId);
                    }
                }
                IFCRebars = parentIds.Select(it => doc.GetElement(it)).ToList();
                #endregion

                #region Поиск опалубки по хосту и геометрии 

                //поиск минимального BoundinBox по IFC стержням          
                List<ElementId> UndefRebarID = new List<ElementId>();
                foreach (Element rebar in IFCRebars)
                {


                    ElementId rebarId = rebar.Id;

                    //формируем контур вокруг с-ва
                    BoundingBoxXYZ box = doc.GetElement(rebarId).get_BoundingBox(null);
                    Outline kontur = new Outline(new XYZ(box.Min.X, box.Min.Y, box.Min.Z),
                                                 new XYZ(box.Max.X, box.Max.Y, box.Max.Z));

                    //ищем кандидата на пересечение
                    BoundingBoxIntersectsFilter filterColl = new BoundingBoxIntersectsFilter(kontur);   //фильтр на элементы которые пересекаются с контуром ifc-арматуры
                    FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);  //фильтр на элементы которые видны на виде

                    //список категорий возможных кандидатов
                    List<BuiltInCategory> categories = new List<BuiltInCategory>
                    {
                        BuiltInCategory.OST_Walls,                  //стены
                        BuiltInCategory.OST_StructuralFraming,      // балки
                        BuiltInCategory.OST_Floors,                 // перекрытия
                        BuiltInCategory.OST_StructuralColumns,      //несущие колонны
                        BuiltInCategory.OST_StructuralFoundation,   //фундаменты
                        BuiltInCategory.OST_Stairs,                 //лестницы
                    };
                    ElementMulticategoryFilter filterCat = new ElementMulticategoryFilter(categories);  //
                                                                                                        //FilteredElementCollector candidates = new FilteredElementCollector(doc);    //фильтр на все элементы в файле

                    //список исключений
                    ICollection<ElementId> idsExclude = new List<ElementId>();
                    idsExclude.Add(rebarId);

                    //формируем список кандидатов для опалубки-родителя
                    List<Element> collElem = new List<Element>();
                    List<ElementId> collElemID = new List<ElementId>();

                    if (rebar.Category != null && rebar.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                    {
                        collElem.Add((rebar as FamilyInstance).Host);   //если это rebar с-во окна, то возвращает его хост 
                    }
                    else
                    {
                        //элементы видимые на виде, исключая текущий стержень/с-во, из списка разрешенных категорий, пересекающие BoundingBox арматуры
                        collElem = collector.Excluding(idsExclude).WherePasses(filterCat).WherePasses(filterColl).ToElements().ToList();
                    }

                    if (collElem.Count > 1)
                    {

                        //элемент с максимальным объемом пересечения
                        Element hostElement = null;
                        //переменная для хранения макисмального объема пересечения
                        double tempIntersectSolidVolume = 0;

                        //Солиды арматуры
                        List<Solid> rebarSolids = GetElementSolids(rebar);
                        // Игнор родительских пустых семейств
                        if (rebarSolids.All(s => s.Volume == 0)) break;


                        foreach (Element elem in collElem)
                        {
                            //Солиды опалубок
                            List<Solid> elemSolids = GetElementSolids(elem);
                            // Игнор пустых основ
                            if (elemSolids.All(s => s.Volume == 0)) break;

                            try
                            {
                                foreach (Solid rebarSolid in rebarSolids)
                                {
                                    foreach (Solid elemSolid in elemSolids)
                                    {
                                        Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(rebarSolid, elemSolid, BooleanOperationsType.Intersect);
                                        if (intersectionSolid != null && intersectionSolid.Volume > tempIntersectSolidVolume)
                                        {
                                            tempIntersectSolidVolume = intersectionSolid.Volume;
                                            hostElement = elem;
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

                        collElem.Clear();
                        collElem.Add(hostElement);

                    }
                    rebarToStruct[rebar] = new List<Element>(collElem);
                    collElemID = collElem.Select(a => a.Id).ToList();
                }
                #endregion

                #region Транзакция
                using (Transaction tr = new Transaction(doc, "Маркировка IFC-арматуры"))
                {
                    tr.Start();
                    foreach (var pair in rebarToStruct)
                    {
                        Element rebar = pair.Key;
                        List<Element> structElems = pair.Value;
                        if (structElems.Count == 1)
                        {
                            Parameter param = rebar.LookupParameter("Мрк.МаркаКонструкции");
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(structElems[0].get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString());
                            }
                        }
                        else
                        {
                            if (structElems == null || structElems.Count == 0)
                                UndefRebarID.Add(rebar.Id);
                            //Parameter param = rebar.LookupParameter("Мрк.МаркаКонструкции");
                        }
                    }
                    tr.Commit(); // или trans.Rollback() при необходимости
                }
                #endregion

                sw.Stop();
                Print("Время выполнения: " + sw.Elapsed + "\n\nНеобработанные стержни будут выделены в модели");
                Selection sel = uidoc.Selection;
                //выделяет непромаркированные стержни
                sel.SetElementIds(UndefRebarID);
            }
            catch (Exception e)
            {
                TaskDialog.Show("Ошибка", e.Message);
            }
            return Result.Succeeded;
        }


        void Print(string a)
        {
            TaskDialog.Show("Итог", a);
        }

        private List<Solid> GetElementSolids(Element element)
        {
            List<Solid> solidColl = new List<Solid>();

            Options opt = new Options() { DetailLevel = ViewDetailLevel.Fine };
            opt.ComputeReferences = true;
            GeometryElement geomElem = element.get_Geometry(opt);
            if (geomElem != null)
            {
                GetSolidsFromGeomElem(geomElem, Transform.Identity, solidColl);
            }

            return solidColl.Where(s => s.Volume > 0).ToList();
        }

        private void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solids)
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



    }
}
