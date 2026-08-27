using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Loader.Common;
using KPLN_MEPBender.Common;
using KPLN_MEPBender.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.ExecutableCommand
{
    public sealed class SelectRouteElementsExcCmd : IExecutableCommand
    {
        private readonly MepBenderM _entity;

        public SelectRouteElementsExcCmd(MepBenderM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            try
            {
                UIDocument uiDoc = app.ActiveUIDocument;
                if (uiDoc == null)
                    return Result.Cancelled;

                IList<Reference> references = uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    new MepBenderSelectionFilter(uiDoc.Document, MepBenderSelectionMode.Route),
                    "Выбери трубы, воздуховоды или кабельные лотки");

                _entity.SetRoutes(references.Select(r => r.ElementId));
                _entity.SetStatus(null, $"Трасс для изменения: {_entity.RouteElementIds.Count}");

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                _entity.SetStatus(null, "Выбор трасс отменён.");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("MEP Bender", ex.ToString());
                _entity.SetStatus("Ошибка при выборе трасс.");
                return Result.Failed;
            }
        }
    }
}
