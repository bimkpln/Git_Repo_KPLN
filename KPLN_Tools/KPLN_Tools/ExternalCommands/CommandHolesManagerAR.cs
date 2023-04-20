using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.HolesManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandHolesManagerAR : IExternalCommand
    {
        private readonly List<BuiltInCategory> _builtInCategories = new List<BuiltInCategory>()
        { 
            // ОВВК
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_MechanicalEquipment,
            // ЭОМСС
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region  Получаю и проверяю коллекцию классов-отверстий в стенах
            IEnumerable<FamilyInstance> holesElems = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                .Where(e => e.Symbol.FamilyName.StartsWith("199_Отверстие"));
            #endregion

            #region  Обрабатываю отверстия
            try
            {
                CheckFamilies(holesElems);

                IEnumerable<RevitLinkInstance> linkedModels = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Where(lm => !(lm.Name.Contains("_KR_") | lm.Name.Contains("_КР_")))
                    .Cast<RevitLinkInstance>();

                foreach (RevitLinkInstance linkedModel in linkedModels)
                {
                    ARLinkData linkData = new ARLinkData(linkedModel);
                    linkData.SetFEC(_builtInCategories);

                    foreach (FamilyInstance fi in holesElems)
                    {
                        BoundingBoxXYZ fiBBox = HoleGeometry(fi);
                        Transform transform = linkData.CurrentLink.GetTransform();

                        BoundingBoxIntersectsFilter filter = CreateFilter(fiBBox, transform);
                        foreach (FilteredElementCollector fec in linkData.CurrentFEC)
                        {
                            IEnumerable<Element> solidIntesectFiIOSElems = fec
                                .WherePasses(filter)
                                .Where(x => !(x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("501_")))
                                .Cast<Element>();
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.InnerException.Message}", KPLN_Loader.Preferences.MessageType.Header);
                else
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.Message}", KPLN_Loader.Preferences.MessageType.Header);

                return Result.Cancelled;
            }
            #endregion

            return Result.Succeeded;
        }

        /// <summary>
        /// Проверка отверстий перед запуском
        /// </summary>
        /// <returns></returns>
        private void CheckFamilies(IEnumerable<FamilyInstance> elems)
        {
            if (!(elems.Any()))
                throw new Exception("Не удалось определить семейства. Поиск осуществялется по категории 'Оборудование', и имени, которое начинается с '199_Отверстие'");
        }

        /// <summary>
        /// Проверка получение геометрии
        /// </summary>
        private BoundingBoxXYZ HoleGeometry(Element elem)
        {
            BoundingBoxXYZ result = null;

            GeometryElement geomElem = elem
                .get_Geometry(new Options()
                {
                    // Особенность семейства, что его геометрия присутсвует ТОЛЬКО на низком уровне детализации
                    DetailLevel = ViewDetailLevel.Coarse,
                });

            foreach (GeometryInstance inst in geomElem)
            {
                Transform transform = inst.Transform;

                GeometryElement instGeomElem = inst.GetInstanceGeometry();
                foreach (GeometryObject obj in instGeomElem)
                {
                    Solid solid = obj as Solid;
                    if (solid != null && solid.Volume != 0)
                    {
                        BoundingBoxXYZ bbox = solid.GetBoundingBox();
                        bbox.Transform = transform;
                        result = new BoundingBoxXYZ()
                        {
                            Max = transform.OfPoint(bbox.Max),
                            Min = transform.OfPoint(bbox.Min),
                        };

                        return result;
                    }
                }
            }

            throw new Exception($"Не удалось определить Solid у элемента с id {elem.Id}. Обратись к разработчику!");
        }

        /// <summary>
        /// Создание фильтра, для поиска элементов, с которыми пересекается BoundingBoxXYZ
        /// </summary>
        private static BoundingBoxIntersectsFilter CreateFilter(BoundingBoxXYZ bbox, Transform transform)
        {
            double minX = bbox.Min.X;
            double minY = bbox.Min.Y;

            double maxX = bbox.Max.X;
            double maxY = bbox.Max.Y;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);

            Outline outline = new Outline(
                transform.OfPoint(new XYZ(sminX, sminY, bbox.Min.Z - 50)), 
                transform.OfPoint(new XYZ(smaxX, smaxY, bbox.Max.Z + 50)));

            return new BoundingBoxIntersectsFilter(outline);
        }
    }
}
