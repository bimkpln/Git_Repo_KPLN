using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Loader.Common;
using KPLN_MEPBender.Common;
using KPLN_MEPBender.Forms.Entities;
using KPLN_MEPBender.Services.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.ExecutableCommand
{
    public sealed class SelectObstaclesExcCmd : IExecutableCommand
    {
        private readonly MepBenderM _entity;
        private readonly MepBenderObstacleSelectionSource _source;

        public SelectObstaclesExcCmd(MepBenderM entity, MepBenderObstacleSelectionSource source)
        {
            _entity = entity;
            _source = source;
        }

        public Result Execute(UIApplication app)
        {
            try
            {
                UIDocument uiDoc = app.ActiveUIDocument;
                if (uiDoc == null)
                    return Result.Cancelled;

                ObjectType objectType = _source == MepBenderObstacleSelectionSource.Link
                    ? ObjectType.LinkedElement
                    : ObjectType.Element;

                IList<Reference> references = uiDoc.Selection.PickObjects(
                    objectType,
                    new MepBenderSelectionFilter(uiDoc.Document, MepBenderSelectionMode.Obstacle, _source),
                    _source == MepBenderObstacleSelectionSource.Link
                        ? "Выбери элементы-препятствия из связанной модели"
                        : "Выбери элементы-препятствия из текущей модели");

                _entity.AddObstacles(references
                    .Select(r => LinkTransformHelper.CreateReference(uiDoc.Document, r))
                    .Where(r => r != null));
                _entity.SetStatus(null, $"Огибаемых элементов: {_entity.ObstacleReferences.Count}");

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                _entity.SetStatus(null, "Выбор огибаемых элементов отменён.");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("MEP Bender", ex.ToString());
                _entity.SetStatus("Ошибка при выборе огибаемых элементов.");
                return Result.Failed;
            }
        }
    }
}
