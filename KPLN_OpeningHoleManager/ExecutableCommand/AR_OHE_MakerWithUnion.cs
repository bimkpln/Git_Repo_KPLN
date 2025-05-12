using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Core.MainEntity;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по расстановке и чистке отверстий АР С ОБЪЕДИНЕНИЕМ
    /// </summary>
    internal sealed class AR_OHE_MakerWithUnion : IExecutableCommand
    {
        private readonly string _transName;
        private readonly AROpeningHoleEntity[] _arEntities;
        private readonly ViewModel _viewModel;
        
        /// <summary>
        /// Конструктор для обработки исходной коллекции, чтобы далее использовать полученный результат
        /// </summary>
        /// <param name="arEntities"></param>
        public AR_OHE_MakerWithUnion(AROpeningHoleEntity[] arEntities, string transName, ViewModel viewModel)
        {
            _arEntities = arEntities;
            _transName = transName;
            _viewModel = viewModel;
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;

            bool creationCanceledBySize = false;
            using (Transaction trans = new Transaction(doc, _transName))
            {
                trans.Start();


                // Создаю новые элементы
                foreach (AROpeningHoleEntity arEntity in _arEntities)
                {
                    arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc, arEntity.AR_OHE_HostElement);
                }

                // Делаю проверку на возможность объединения
                Dictionary<int, List<AROpeningHoleEntity>> arEntitiesForUnionByGroups = GetAROpeningsFromIOSElems_UnionByParams(_arEntities);
                if (arEntitiesForUnionByGroups.Any())
                {
                    List<AROpeningHoleEntity> arEntitiesForUnion = new List<AROpeningHoleEntity>();
                    List<AROpeningHoleEntity> arEntitiesForDelete = new List<AROpeningHoleEntity>();
                    foreach (var kvp in arEntitiesForUnionByGroups)
                    {
                        arEntitiesForDelete.AddRange(kvp.Value);
                        AROpeningHoleEntity unionOHE = AROpeningHoleEntity.CreateUnionOpeningHole(doc, kvp.Value.ToArray(), false);

                        // Уточняю параметры исходя из материала основы
#if Debug2020 || Revit2020
                        double minHeight = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS);
                        double minWidht = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS);
                        double minRadius = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS)/ 2;
                        if (unionOHE.AR_OHE_IsHostElementKR)
                        {
                            minHeight = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS);
                            minWidht = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS);
                            minRadius = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, DisplayUnitType.DUT_MILLIMETERS) / 2;
                        }
#else
                        double minHeight = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
                        double minWidht = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
                        double minRadius = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")) / 2;
                        if (unionOHE.AR_OHE_IsHostElementKR)
                        {
                            minHeight = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
                            minWidht = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
                            minRadius = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinHeightValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")) / 2;
                        }
#endif

                        if (unionOHE.OHE_Height >= minHeight
                            || unionOHE.OHE_Width >= minWidht
                            || unionOHE.OHE_Radius >= minRadius)
                            arEntitiesForUnion.Add(unionOHE);
                        else
                            creationCanceledBySize = true;
                    }


                    // Сначала удаляю отверстия, из которых состоит объединённое
                    foreach (AROpeningHoleEntity arEntity in arEntitiesForDelete)
                    {
                        // Может быть, что элемента ещё нет в файле
                        if (arEntity.OHE_Element.IsValidObject)
                            doc.Delete(arEntity.OHE_Element.Id);
                    }


                    // Затем создаю новые элементы
                    foreach (AROpeningHoleEntity arEntity in arEntitiesForUnion)
                    {
                        arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc, arEntity.AR_OHE_HostElement);
                    }
                }

                trans.Commit();

                if (creationCanceledBySize)
                {
                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                        MainInstruction = $"При расстановке отверстий часть отверстий НЕ создано из-за того, что они попали в допуск по размерам (т.е. их создавать не нужно)",
                    }.Show();
                }

            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию отверстий АР, которые могут быть объеденены по нужным параметрам. 
        /// где: int - id любого отверстия группы, вокруг которого будет вестись объединение, List<AROpeningHoleEntity> - коллекция для объединения
        /// </summary>
        /// <returns></returns>
        private Dictionary<int, List<AROpeningHoleEntity>> GetAROpeningsFromIOSElems_UnionByParams(AROpeningHoleEntity[] arEntities)
        {
            var adjacency = new Dictionary<int, List<int>>();
            var idToEntity = arEntities.ToDictionary(e => e.OHE_Element.Id.IntegerValue, e => e);
#if Debug2020 || Revit2020
            double ar_minDistance = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinDistanceValue, DisplayUnitType.DUT_MILLIMETERS);
            double kr_minDistance = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinDistanceValue, DisplayUnitType.DUT_MILLIMETERS);
#else
            double ar_minDistance = UnitUtils.ConvertToInternalUnits(_viewModel.AR_OpenHoleMinDistanceValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
            double kr_minDistance = UnitUtils.ConvertToInternalUnits(_viewModel.KR_OpenHoleMinDistanceValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif

            // Ствараем граф
            foreach (var ent1 in arEntities)
            {
                int id1 = ent1.OHE_Element.Id.IntegerValue;
                double minDistance = ent1.AR_OHE_IsHostElementKR ? kr_minDistance : ar_minDistance;

                foreach (var ent2 in arEntities)
                {
                    int id2 = ent2.OHE_Element.Id.IntegerValue;
                    if (id1 == id2) continue;

                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(ent1.OHE_Solid, ent2.OHE_Solid, BooleanOperationsType.Intersect);
                    bool isConnected = intersectSolid != null && intersectSolid.Volume > 0;

                    if (!isConnected)
                    {
                        double dist = GeometryWorker.GetMinimumDistanceBetweenSolids(ent1.OHE_Solid, ent2.OHE_Solid);
                        isConnected = dist <= minDistance;
                    }

                    if (isConnected)
                    {
                        if (!adjacency.ContainsKey(id1)) adjacency[id1] = new List<int>();
                        if (!adjacency[id1].Contains(id2)) adjacency[id1].Add(id2);

                        if (!adjacency.ContainsKey(id2)) adjacency[id2] = new List<int>();
                        if (!adjacency[id2].Contains(id1)) adjacency[id2].Add(id1);
                    }
                }
            }

            // Знаходзім усе групы праз BFS
            var visited = new HashSet<int>();
            var result = new Dictionary<int, List<AROpeningHoleEntity>>();

            foreach (int startId in idToEntity.Keys)
            {
                if (visited.Contains(startId)) continue;

                var queue = new Queue<int>();
                var group = new List<AROpeningHoleEntity>();

                queue.Enqueue(startId);
                visited.Add(startId);

                while (queue.Count > 0)
                {
                    int currentId = queue.Dequeue();
                    group.Add(idToEntity[currentId]);

                    if (!adjacency.ContainsKey(currentId)) continue;

                    foreach (int neighborId in adjacency[currentId])
                    {
                        if (!visited.Contains(neighborId))
                        {
                            visited.Add(neighborId);
                            queue.Enqueue(neighborId);
                        }
                    }
                }

                // Выкарыстоўваем першы ID у групе як ключ
                result[group[0].OHE_Element.Id.IntegerValue] = group;
            }

            return result;
        }
    }
}
