using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Views_Ribbon.Common.Lists;
using KPLN_Views_Ribbon.Forms;

namespace KPLN_Views_Ribbon.ExternalCommands.Lists
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    internal class CommandListRename : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;

            //Workin with elements
            List<ViewSheet> mixedSheetsList = new List<ViewSheet>();
            List<Element> falseElemList = new List<Element>();
            List<ElementId> selIds = sel.GetElementIds().ToList();
            foreach (ElementId selId in selIds)
            {
                Element elem = doc.GetElement(selId);
                int catId = elem.Category.Id.IntegerValue;
                if (catId.Equals((int)BuiltInCategory.OST_Sheets))
                {
                    ViewSheet curViewSheet = elem as ViewSheet;
                    mixedSheetsList.Add(curViewSheet);
                }
                else
                {
                    falseElemList.Add(elem);
                }
            }
            List<ViewSheet> sortedSheets = mixedSheetsList.OrderBy(s => s.SheetNumber.ToString()).ToList();

            //Main part of code
            if (mixedSheetsList.Count != 0)
            {
                ParameterSet titleBlockParams = mixedSheetsList[0].Parameters;
                FormListRename inputForm = new FormListRename(titleBlockParams);
                inputForm.ShowDialog();

                using (Transaction trans = new Transaction(doc, "KPLN: Перенумеровать лист"))
                {
                    trans.Start();
                    try
                    {
                        //Меняю номер листа с использованием символов Юникода
                        if ((bool)inputForm.isUNICode.IsChecked && inputForm.IsRun)
                        {
                            UniEntity cmbSelUni = (UniEntity)inputForm.cmbUniCode.SelectedItem;
                            int uniNumber = Int32.Parse(inputForm.strNumUniCode.Text) - 1;
                            bool isRun = true;
                            while (isRun)
                            {
                                uniNumber++;
                                isRun = UseUniCodes(sortedSheets, cmbSelUni, uniNumber, (bool)inputForm.isEditToUni.IsChecked);
                            }
                        }
                    
                        //Меняю номер листа с использованием префиксов
                        else if ((bool)inputForm.isPrefix.IsChecked && inputForm.IsRun)
                        {
                            string userPrefix = inputForm.prfTextBox.Text;
                            userPrefix = userPrefix.Length > 0 ? userPrefix = userPrefix + "/" : userPrefix;
                            int userNumber = Convert.ToInt32(inputForm.strNumTextBox.Text) - 1;
                            if ((bool)inputForm.isRenumbering.IsChecked)
                            {
                                UsePrefix(sortedSheets, userPrefix, userNumber);
                            }
                            else
                            {
                                UsePrefix(sortedSheets, userPrefix);
                            }
                        }

                        //Меняю номер листа без префиксов и Юникодов
                        else if ((bool)inputForm.isClearRenumb.IsChecked && inputForm.IsRun)
                        {
                            int startNumber = Convert.ToInt32(inputForm.strClearNumTextBox.Text);
                            ClearRenumber(sortedSheets, startNumber);
                        }

                        // Заполняю пользовательский параметр для нумерации в штампе
                        else if ((bool)inputForm.isRefreshParam.IsChecked && inputForm.IsRun)
                        {
                            Parameter cmbSelPar = (Parameter)inputForm.cmbParam.SelectedItem;
                            ParamRefresh(sortedSheets, cmbSelPar);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("Sheet number is already in use"))
                        {
                            TaskDialog.Show("Предупреждение", "Такой номер листа уже есть. Работа экстренно завершена!", TaskDialogCommonButtons.Ok);
                            trans.RollBack();
                            return Result.Failed;
                        }
                        else
                        {
                            TaskDialog.Show("Ошибка", $"Отправь в BIM-отдел\n\n{ex.StackTrace}\n{ex.Message}\n\nРабота экстренно завершена!", TaskDialogCommonButtons.Ok);
                            trans.RollBack();
                            return Result.Failed;
                        }
                    }
                    trans.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Ошибка", "В выборке нет ни одного листа, или вообще ничего не выбрано :(", TaskDialogCommonButtons.Ok);
                return Result.Cancelled;
            }
            if (falseElemList.Count != 0)
            {
                int cnt = 0;
                foreach (Element elem in falseElemList)
                {
                    cnt++;
                }
                string msg = string.Format("Если что, были случайно выбраны элементы, которые не являются листами. Успешно проигнорировано {0} штук/-и", cnt.ToString());
                TaskDialog.Show("Предупреждение", msg, TaskDialogCommonButtons.Ok);
            }

            // Обновляю нумерацию в браузере проекта
            DockablePaneId dpId = DockablePanes.BuiltInDockablePanes.ProjectBrowser;
            DockablePane dP = new DockablePane(dpId);
            dP.Show();
            
            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для замены номер листа с использованием символов Юникода
        /// </summary>
        /// <example> 
        /// "7.1" преобразуется в "символЮникода7.1"
        /// </example>
        private bool UseUniCodes(List<ViewSheet> sortedSheets, UniEntity cmbSelUni, int counter, bool isConvertToUni)
        {
            try
            {
                foreach (ViewSheet curVSheet in sortedSheets)
                {
                    string trueNumb;
                    if (isConvertToUni)
                    {
                        string strVSheetNumb = curVSheet.SheetNumber;
                        trueNumb = UserNumber(strVSheetNumb);
                    }
                    else
                    {
                        trueNumb = curVSheet.SheetNumber;
                    }
                    curVSheet.SheetNumber = String.Concat(Enumerable.Repeat(cmbSelUni.Code, counter)) + trueNumb;
                    try
                    {

                        if (counter == 1)
                        {
                            curVSheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e")).Set(cmbSelUni.Name);
                        }
                        else
                        {
                            curVSheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e")).Set($"{counter}X{cmbSelUni.Name}");
                        }
                    }
                    catch { }
                }
                return false;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                return true;
            }
        }

        /// <summary>
        /// Метод для замены номер листа с использованием символов префикса и стартового номера листа
        /// </summary>
        /// <example> 
        /// "7" преобразуется в "АР/8", или в "008"
        /// </example>
        private void UsePrefix(List<ViewSheet> sortedSheets, string userPrefix, int userNumber)
        {
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                userNumber++;
                RenumberWithPrefix(curVSheet, userPrefix, userNumber.ToString());
            }
        }

        /// <summary>
        /// Метод для замены номер листа с использованием символов префикса
        /// </summary>
        /// <example> 
        /// "7.1" преобразуется в "АР/7.1", или в "007.1"
        /// </example>
        private void UsePrefix(List<ViewSheet> sortedSheets, string userPrefix)
        {
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                string strVSheetNumb = curVSheet.SheetNumber;
                string onlyNumb = UserNumber(strVSheetNumb);
                RenumberWithPrefix(curVSheet, userPrefix, onlyNumb);
            }
        }

        /// <summary>
        /// Метод для замены номер листа с использованием стартового номера
        /// </summary>
        /// <example> 
        /// "7" преобразуется в "8"
        /// </example>
        private void ClearRenumber(List<ViewSheet> sortedSheets, int startNumber)
        {
            // Получаю стартовую разницу между номерами
            int deltaNumber;
            string constPartOfStartNumber = sortedSheets[0].SheetNumber.Split('.')[0];
            if (Int32.TryParse(UserNumber(constPartOfStartNumber), out int constStartNumber))
            {
                deltaNumber = Math.Abs(startNumber - constStartNumber);
            }
            else
            {
                throw new Exception("Сначала очисти от приставок, а потом запускай нумерацию");
            }
            
            // Задаю нумерацию с учетом стартовой разницы
            string constPartOfNumber;
            string varPartOfNumber;
            sortedSheets.Reverse();
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                if (curVSheet.SheetNumber.Contains("."))
                {
                    constPartOfNumber = curVSheet.SheetNumber.Split('.')[0];
                    varPartOfNumber = curVSheet.SheetNumber.Split('.').Skip(1).Aggregate((x, y) => x + y);
                    if (Int32.TryParse(UserNumber(constPartOfNumber), out int constNumber))
                    {
                        curVSheet.SheetNumber = (constNumber + deltaNumber).ToString() + "." + varPartOfNumber;
                    }
                    else
                    {
                        throw new Exception("Сначала очисти от приставок, а потом запускай нумерацию");
                    }
                }
                else
                {
                    if (Int32.TryParse(UserNumber(curVSheet.SheetNumber), out int constNumber))
                    {
                        curVSheet.SheetNumber = (constNumber + deltaNumber).ToString();
                    }
                    else
                    {
                        throw new Exception("Сначала очисти от приставок, а потом запускай нумерацию");
                    }
                }
            }
        }

        /// <summary>
        /// Метод для заполнения номером пользовательского параметра
        /// </summary>
        /// <example> 
        /// "АР/7.1" заполниться в выбранный параметр как "7.1"
        /// </example>
        private void ParamRefresh(List<ViewSheet> sortedSheets, Parameter cmbSelPar)
        {
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                string strVSheetNumb = curVSheet.SheetNumber;
                string cmbSelParName = cmbSelPar.Definition.Name;
                string onlyNumb = UserNumber(strVSheetNumb);
                if (Int32.TryParse(strVSheetNumb, out int refreshNumber))
                {
                    curVSheet.LookupParameter(cmbSelParName).Set(refreshNumber.ToString());
                }
                else
                {
                    curVSheet.LookupParameter(cmbSelParName).Set(onlyNumb);
                }
            }
        }

        /// <summary>
        /// Метод для изменения нумерации листов с приставкой.
        /// <example> 
        /// Например: 7.1 преобразуется в АР/007.1, или в 007.1 (в зависимости от наличия приставки)
        /// </example>
        /// </summary>
        /// <param name="sheet">Лист, который нужно перенумеровать</param>
        /// <param name="prefix">Значение приставки к номеру</param>
        /// <param name="number">Номер листа (в формате 1, 2, 3....)</param>
        private void RenumberWithPrefix(ViewSheet sheet, string prefix, string number)
        {
            string ZeroNumber;
            if (OnlyNumber(number) < 10)
            {
                ZeroNumber = $"00{number}";
            }
            else if (OnlyNumber(number) < 100)
            {
                ZeroNumber = $"0{number}";
            }
            else
            {
                ZeroNumber = $"{number}";
            }
            sheet.SheetNumber = prefix + ZeroNumber;
        }

        /// <summary>
        /// Метод преобразования номера с символами в число.
        /// <example> 
        /// Например: АР/007.1 преобразуется в 7.1
        /// </example>
        /// </summary>
        private string UserNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            foreach (char c in charArray)
            {
                if (Char.IsLetter(c))
                {
                    number = number.Trim(c);
                }
                else if (c.Equals('.'))
                {
                    continue;
                }
                else if (Char.IsPunctuation(c))
                {
                    number = number.Trim(c);
                }
            }
            number = number.Length > 1 ? number.TrimStart('0') : number;
            return number;
        }

        /// <summary>
        /// Метод преобразования номера с символами в число без подразделов
        /// </summary>
        /// <example> 
        /// АР/007.1 преобразуется в 7
        /// </example>
        private int OnlyNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            foreach (char c in charArray)
            {
                if (c.Equals('.'))
                {
                    number = number.Split('.')[0];
                    break;
                }
            }
            return Int32.Parse(new String(number.Where(Char.IsDigit).ToArray()));
        }
    }
}
