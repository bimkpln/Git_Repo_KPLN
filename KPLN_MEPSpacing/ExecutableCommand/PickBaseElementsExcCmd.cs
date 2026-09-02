using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Loader.Common;
using KPLN_MEPSpacing.Common;
using KPLN_MEPSpacing.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_MEPSpacing.ExecutableCommand
{
    internal sealed class PickBaseElementsExcCmd : IExecutableCommand
    {
        private readonly MepSpacingM _entity;
        private readonly Window _window;

        public PickBaseElementsExcCmd(MepSpacingM entity, Window window)
        {
            _entity = entity;
            _window = window;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            try
            {
                _window?.Hide();
                IList<Reference> references = uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    new MepCurveSelectionFilter(),
                    "Выбери базовые трубы, воздуховоды или лотки. Они останутся на месте");

                List<ElementId> baseElementIds = references
                    .Select(r => r.ElementId)
                    .Where(id => id != null && !id.Equals(ElementId.InvalidElementId))
                    .GroupBy(GetElementIdValue)
                    .Select(group => group.First())
                    .ToList();

                _entity.SetBaseElementIds(baseElementIds);
                _entity.SetResultStatus($"Базовых элементов выбрано: {baseElementIds.Count}.");
                _entity.UserHelp = "Базовые элементы не будут двигаться. Закреплённые элементы тоже будут считаться неподвижными.";

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                _entity.SetErrorStatus("Не удалось выбрать базовые элементы.");
                _entity.UserHelp = ex.Message;
                return Result.Failed;
            }
            finally
            {
                _window?.Show();
                _window?.Activate();
            }
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }
    }
}
