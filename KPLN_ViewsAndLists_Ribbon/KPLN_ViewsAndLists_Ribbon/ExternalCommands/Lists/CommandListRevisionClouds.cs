using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandListRevisionClouds : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            View activeView = doc.ActiveView;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IEnumerable<ViewSheet> sheets = collector
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(sheet => !sheet.IsTemplate);

            if (!sheets.Any())
            {
                MessageBox.Show($"В проекте нет ни одного листа", "KPLN. Внимание");
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "KPLN. Обновление ИЗМов"))
            {
                trans.Start();

                // Заполнение "Ш.ПримечаниеЛиста"
                foreach (ViewSheet sheet in sheets)
                {
                    List<ElementId> revisionIds = sheet.GetAllRevisionIds().ToList();
                    if (revisionIds.Count > 0)
                    {
                        Parameter statusParam = sheet.LookupParameter("Ш.ШифрСтатусЛиста");

                        if (statusParam != null)
                        {
                            int code = statusParam.AsInteger();
                            string fieldValue = GetRevisionStatusString(code);

                            Parameter targetParam = sheet.LookupParameter("Ш.ПримечаниеЛиста");
                            if (targetParam != null && targetParam.StorageType == StorageType.String)
                            {
                                targetParam.Set(fieldValue);
                            }

                            if (string.IsNullOrWhiteSpace(fieldValue))
                            {
                                string result = "";

                                for (int i = 0; i < revisionIds.Count; i++)
                                {
                                    result += $"Изм. {i + 1}";
                                    if (i < revisionIds.Count - 1)
                                        result += ", ";
                                }

                                if (!string.IsNullOrWhiteSpace(result))
                                {
                                    targetParam.Set(result);
                                }
                            }
                        }
                    }
                }

                // Заполнение "Кол.уч."
                foreach (ViewSheet sheet in sheets)
                {
                    string[] cloudsCounts = GetCloudsCountOnSheet(sheet);

                    sheet.LookupParameter("Ш.КолвоУч1Текст").Set("");
                    sheet.LookupParameter("Ш.КолвоУч2Текст").Set("");
                    sheet.LookupParameter("Ш.КолвоУч3Текст").Set("");
                    sheet.LookupParameter("Ш.КолвоУч4Текст").Set("");

                    for (int i = 0; i < cloudsCounts.Length; i++)
                    {
                        string val = cloudsCounts[i];
                        string paramName = "Ш.КолвоУч" + (i + 1).ToString() + "Текст";
                        sheet.LookupParameter(paramName).Set(val);
                    }                                                                    
                }

                trans.Commit();
            }
            
            MessageBox.Show($"Данные об ИЗМах обновлены", "KPLN. Информация");
            return Result.Succeeded;
        }


        // Получение содержимого "Ш.ШифрСтатусЛиста"
        private string GetRevisionStatusString(int code)
        {
            var result = new List<string>();
            string codeStr = code.ToString().PadLeft(4, '0');

            for (int i = 0; i < codeStr.Length && i < 4; i++)
            {
                char c = codeStr[i];
                string label = null;

                if (c == '2') label = "Зам.";
                else if (c == '3') label = "Нов.";

                if (label != null)
                {
                    int revNum = i + 1;
                    result.Add($"Изм. {revNum}({label})");
                }
            }

            return string.Join(", ", result);
        }

        // Считаем "облачка"
        private string[] GetCloudsCountOnSheet(ViewSheet sheet)
        {
            Document doc = sheet.Document;
            List<RevisionCloud> clouds = new FilteredElementCollector(doc)
                .OfClass(typeof(RevisionCloud))
                .Cast<RevisionCloud>()
                .Where(i => i.OwnerViewId == sheet.Id)
                .ToList();

            List<Revision> allRevisionsOnSheet = sheet.GetAllRevisionIds()
                .Select(id => doc.GetElement(id))
                .OfType<Revision>()
                .OrderBy(r => r.IssuedTo) 
                .ToList();

            var allRevisionsGrouped = allRevisionsOnSheet
                .GroupBy(r => r.IssuedTo)
                .ToDictionary(g => g.Key, g => g.ToList());

            Dictionary<string, List<Revision>> revsBase = new Dictionary<string, List<Revision>>();
            foreach (Revision rev in allRevisionsOnSheet)
            {
                string utvDlya = rev.IssuedTo;
                if (!revsBase.ContainsKey(utvDlya))
                    revsBase[utvDlya] = new List<Revision>();
                revsBase[utvDlya].Add(rev);
            }

            int revisionsCount = allRevisionsOnSheet.Count;
            List<Revision> lastRevisionsOnSheet = new List<Revision>();

            if (revisionsCount > 4)
            {
                lastRevisionsOnSheet = allRevisionsOnSheet.GetRange(revisionsCount - 4, 4);
            }
            else
            {
                lastRevisionsOnSheet = allRevisionsOnSheet;
            }

            int revsCount = lastRevisionsOnSheet.Count;
            string[] cloudsCounts = new string[revsCount];

            for (int i = 0; i < revsCount; i++)
            {
                Revision rev = lastRevisionsOnSheet[i];
                int curRevCloudsCount = clouds.Where(c => c.RevisionId == rev.Id).ToList().Count;
                string val = "";
                if (curRevCloudsCount == 0) val = "-";
                else val = curRevCloudsCount.ToString();

                cloudsCounts[i] = val;
            }
            return cloudsCounts;
        }
    }
}     