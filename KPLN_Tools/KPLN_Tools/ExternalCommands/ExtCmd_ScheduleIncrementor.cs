using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_ScheduleIncrementor : IExternalCommand
    {
        internal const string PluginName = "ОВВК: Нумерация";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            View activeView = doc.ActiveView;
            if (!(activeView is ViewSchedule viewSchedule))
            {
                MessageBox.Show("Открой нужную спецификацию для анализа", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return Result.Cancelled;
            }

            try
            {
                // Информация по скрытым столбцам
                ScheduleDefinition def = viewSchedule.Definition;
                List<ScheduleFieldId> hiddenFieldIds = new List<ScheduleFieldId>();
                int fieldCount = def.GetFieldCount();
                for (int i = 0; i < fieldCount; i++)
                {
                    ScheduleField field = def.GetField(i);
                    if (field.IsHidden) 
                        hiddenFieldIds.Add(field.FieldId);
                }


                // Создаём сущности для анализа
                ScheduleEntity se = new ScheduleEntity()
                {
                    SE_ViewSchedule = viewSchedule,
                    SE_HiddenFieldIds = hiddenFieldIds
                };


                // Открываю все скрытые столбцы в спеке
                if (se.SE_HiddenFieldIds.Count > 0)
                {
                    using (Transaction t = new Transaction(doc, "KPLN: Спецификация показать скрытые"))
                    {
                        t.Start();

                        foreach (ScheduleFieldId fieldId in se.SE_HiddenFieldIds)
                        {
                            if (!def.IsValidFieldId(fieldId))
                                continue;

                            ScheduleField field = def.GetField(fieldId);
                            field.IsHidden = false; 
                        }

                        viewSchedule.RefreshData(); 
                        t.Commit();
                    }
                }

                // Читаю таблицу
                var model = ScheduleHelper.ReadSchedule(se);
                var window = new ScheduleMainForm(model, viewSchedule.Name);


                // Привязка к окну ревит
                var helper = new WindowInteropHelper(window) { Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle };


                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }
    }
}
