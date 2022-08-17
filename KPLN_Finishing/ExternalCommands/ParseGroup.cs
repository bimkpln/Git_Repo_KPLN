using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using KPLN_Loader.Output;
using static KPLN_Finishing.Tools;
using static KPLN_Loader.Output.Output;
using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Finishing.Forms;
using static KPLN_Loader.Preferences;
using KPLN_Finishing.CommandTools;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class ParseGroup : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (Preferences.ShowGroupParserToolTip)
            {
                TaskDialog td = new TaskDialog("Типовые этажи");
                td.TitleAutoPrefix = false;
                td.MainContent = "Краткая инструкция";
                td.FooterText = "Алгоритмы работы:\n" +
                    "1) выбрать группу* с отделкой\n\n" +
                    "* для правильной работы скрипта необходимо иметь типовые этажи и сгруппированные по этажам элементы отделки.";
                td.CommonButtons = TaskDialogCommonButtons.Ok;
                td.VerificationText = "Больше не показывать";
                TaskDialogResult result = td.Show();
                if (td.WasVerificationChecked())
                { Preferences.ShowGroupParserToolTip = false; }
            }
            try
            {
                Group group = null;
                try
                {
                    Reference reference = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element, new GroupSelectionFilter(), "Выберите элемент : <Группа>");
                    group = commandData.Application.ActiveUIDocument.Document.GetElement(reference) as Group;
                }
                catch (Exception)
                { }
                if (group != null)
                {
                    GroupParser.ParseGroup(commandData.Application.ActiveUIDocument.Document, group);
                }
                return Result.Succeeded;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
    }
}