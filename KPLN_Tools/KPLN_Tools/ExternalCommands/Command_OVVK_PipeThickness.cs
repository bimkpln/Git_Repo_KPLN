using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OVVK_PipeThickness : IExternalCommand
    {
        /// <summary>
        /// Параметр толщины стенки труб
        /// </summary>
        private readonly Guid _thicknessParam = new Guid("381b467b-3518-42bb-b183-35169c9bdfb3");
#if Revit2020
        private readonly DisplayUnitType _millimetersRevitType = DisplayUnitType.DUT_MILLIMETERS;
#endif
#if Revit2023 || Debug
        private readonly ForgeTypeId _millimetersRevitType = new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1");
#endif

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            // Получаем используемые диаметры труб
            Pipe[] pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToArray();

            Dictionary<ElementId, List<double>> pipeDict = new Dictionary<ElementId, List<double>>();
            foreach (Pipe pipe in pipes)
            {
                ElementId typeId = pipe.PipeType.Id;
                if (!pipeDict.ContainsKey(typeId))
                    pipeDict.Add(typeId, new List<double>());

                Parameter diameterParameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (diameterParameter != null && diameterParameter.HasValue)
                {
                    double diameter = Math.Round(UnitUtils.ConvertFromInternalUnits(diameterParameter.AsDouble(), _millimetersRevitType), 1);
                    if (pipeDict.TryGetValue(typeId, out List<double> diams) && !diams.Contains(diameter))
                        pipeDict[typeId].Add(diameter);
                }
            }

            // Выводим окно пользователю
            OVVK_PipeThicknessForm mainForm = new OVVK_PipeThicknessForm(doc, pipeDict);
            mainForm.ShowDialog();

            if (mainForm.IsRun)
            {
                using (Transaction trans = new Transaction(doc, "KPLN: Толщина стенки труб"))
                {
                    trans.Start();

                    foreach (Pipe pipe in pipes)
                    {
                        Parameter thicknessParameter = pipe.get_Parameter(_thicknessParam)
                            ?? throw new Exception("В проекте у труб нет параметра \"КП_И_Толщина стенки\", либо он не параметр экземпляра. Добавь его и повтори запуск");

                        PipeType pipeType = pipe.PipeType;
                        ElementId pipeTypeId = pipeType.Id;
                        Parameter diameterParameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParameter != null && diameterParameter.HasValue)
                        {
                            double diameter = Math.Round(UnitUtils.ConvertFromInternalUnits(diameterParameter.AsDouble(), _millimetersRevitType), 1);
                            foreach (PipeThicknessEntity entity in mainForm.PipeThicknessEntities)
                            {
                                if (pipeTypeId.Equals(entity.CurrentPipeType.Id))
                                {
                                    double currentThickness = GetThicknessByPipeThicknessEntityAndDiam(entity, diameter);
                                    thicknessParameter.Set(UnitUtils.ConvertToInternalUnits(currentThickness, _millimetersRevitType));
                                    break;
                                }
                            }
                        }
                    }

                    trans.Commit();

                    return Result.Succeeded;
                }
            }

            return Result.Cancelled;
        }

        private double GetThicknessByPipeThicknessEntityAndDiam(PipeThicknessEntity entity, double diam)
        {
            foreach (PipeTypeDiamAndThickness diamAndThick in entity.CurrentPipeTypeDiamAndThickness)
            {
                if (diam == diamAndThick.CurrentDiameter)
                    return diamAndThick.CurrentThickness;
            }

            return 0;
        }
    }
}
