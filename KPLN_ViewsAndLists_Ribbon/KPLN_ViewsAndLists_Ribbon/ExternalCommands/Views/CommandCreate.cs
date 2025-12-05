using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Views.CsvReaders;
using KPLN_ViewsAndLists_Ribbon.Views.FilterUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    class CommandCreate : IExternalCommand
    {
        [Obsolete]
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                string dllPath = Path.GetDirectoryName(Module._assemblyPath);

                OpenFileDialog openCsvDialog = new OpenFileDialog
                {
                    Filter = "CSV file|*.csv",
                    Title = "Выберите файл CSV (v2018.10.17)",
                    Multiselect = false,
                    InitialDirectory = dllPath
                };

                if (openCsvDialog.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled; ;

                //считываю файл
                string path = openCsvDialog.FileName;
                List<string[]> data = ReadDataFromCSV.Read(path);

                string msg = "";
                int filterCount = 0;

                //одна строка в файле - один фильтр
                foreach (string[] line in data)
                {
                    FilterSourceInfo filterSource = new FilterSourceInfo(line);
                    string filterName = filterSource.FilterName;

                    //Добавляю категории
                    List<ElementId> catIds = new List<ElementId>();
                    foreach (string stringCat in filterSource.Categories)
                    {
                        BuiltInCategory cat = GetBuiltinCategory.GetCategoryByRussianName(stringCat);
                        catIds.Add(new ElementId(cat));
                    }


                    //добавляю критерии фильтрации
                    List<FilterRule> filterRules = new List<FilterRule>();

                    foreach (string[] sourceRule in filterSource.SourceRules)
                    {
                        string paramName = sourceRule[0];
                        string function = sourceRule[1];
                        string value = sourceRule[2];


                        BuiltInCategory cat = GetBuiltinCategory.GetCategoryByRussianName(filterSource.Categories[0]);
                        if (cat == BuiltInCategory.OST_Sections || cat == BuiltInCategory.OST_Elev || cat == BuiltInCategory.OST_Callouts)
                            cat = BuiltInCategory.OST_Views;

                        FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(cat);

                        Parameter param = null;
                        try
                        {
                            foreach (Element elem in collector)
                            {
                                param = elem.LookupParameter(paramName);
                                if (param == null)
                                    continue;
                                break;
                            }
                        }
                        catch { }

                        if (collector.Count() == 0 || param == null)
                        {
                            message = "Ошибка при создании фильтра: " + filterName;
                            message += "\nУстановите как минимум один элемент в категории: " + filterSource.Categories[0];
                            message += "\nТребуемый параметр: " + paramName;
                            return Result.Failed;
                        }

                        FilterRule rule = FilterCreator.CreateRule2(param, function, value);
                        filterRules.Add(rule);
                    }

                    try
                    {
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Создание фильтра" + filterName);
                            ParameterFilterElement filter = ParameterFilterElement.Create(doc, filterName, catIds);
#if R2017 || R2018
                        
                        filter.SetRules(filterRules);
#else
                            ElementParameterFilter epf = new ElementParameterFilter(filterRules);
                            filter.SetElementFilter(epf);
#endif
                            filterCount++;
                            t.Commit();
                        }
                    }
                    catch
                    {
                        msg += filterName + "\n";
                    }

                }
                string finalMessage = "Создано фильтров: " + filterCount.ToString() + "\n";
                if (msg.Length != 0)
                {
                    finalMessage += "Не удалось создать: \n" + msg;
                }

                TaskDialog.Show("Batch filter create", finalMessage);
                return Result.Succeeded;
            }

            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");
                return Result.Failed;
            }
        }
    }
}
