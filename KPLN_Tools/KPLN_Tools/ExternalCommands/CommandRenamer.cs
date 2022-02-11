using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Forms;


namespace KPLN_Tools.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class CommandRenamer : IExternalCommand
    {
        static void SetNameNumber(ViewSheet sheet, string prefix, int number)
        {
            string sheetNumber;
            try
            {
                if (number < 10)
                {
                    sheetNumber = string.Format("00{0}", number.ToString());
                }
                else if (number < 100)
                {
                    sheetNumber = string.Format("0{0}", number.ToString());
                }
                else
                {
                    sheetNumber = string.Format("{0}", number.ToString());
                }
                try
                {
                    sheet.SheetNumber = prefix + sheetNumber;
                }
                catch (Exception) { throw; };


            }
            catch (Exception e)
            {
                TaskDialog.Show("Ошибка", string.Format("Для элемента {0} - Ошибка: {1}", sheet, e));
            }
        }
        static void SetName(ViewSheet sheet, string prefix)
        {
            string beginNumb = sheet.SheetNumber;
            try
            {
                sheet.SheetNumber = prefix + beginNumb;
            }
            catch (Exception e)
            {
                TaskDialog.Show("Ошибка", string.Format("Для элемента {0} - Ошибка: {1}", sheet, e));
            }
        }
        static string GetSliceNumber(int number)
        {
            string strNumb = number.ToString();
            if (strNumb.StartsWith("0"))
            {
                strNumb = strNumb.Remove(0,1);
            }
            if (strNumb.StartsWith("0"))
            {
                strNumb = strNumb.Remove(0, 1);
            }
            return strNumb;
        }

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
                int a = elem.Category.Id.IntegerValue;
                if (elem.Category.Id.IntegerValue == -2003100)
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
                FormRenamer inputForm = new FormRenamer(titleBlockParams);
                inputForm.ShowDialog();
                string prfTxt = inputForm.prfTxt;
                int intNumb = Convert.ToInt32(inputForm.strNumb) - 1;
                using (Transaction trans = new Transaction(doc, "KPLN: Перенумеровать лист"))
                {
                    trans.Start();
                    foreach (ViewSheet curVSheet in sortedSheets)
                    {
                        //Change number in sheet
                        string strVSheetNumb = curVSheet.SheetNumber;
                        if ((bool)inputForm.isRenumbering.IsChecked)
                        { 
                            intNumb++;
                            SetNameNumber(curVSheet, prfTxt, intNumb);
                        }
                        else
                        {
                            try
                            {
                                intNumb = Convert.ToInt32(strVSheetNumb);
                            }
                            catch (Exception)
                            {
                                string onlyNumb = "";
                                foreach(char i in strVSheetNumb)
                                {
                                    if (i < '0' || i > '9')
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        onlyNumb += i;
                                    }
                                }
                                intNumb = Convert.ToInt32(onlyNumb);
                            }
                            
                            if((bool)inputForm.isChangingList.IsChecked)
                            {
                                Parameter cmbSelPar = (Parameter)inputForm.cmbParam.SelectedItem;
                                string cmbSelParName = cmbSelPar.Definition.Name;
                                string sliceNumb = GetSliceNumber(intNumb);
                                curVSheet.LookupParameter(cmbSelParName).Set(sliceNumb);
                            }
                            SetName(curVSheet, prfTxt);
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

            return Result.Succeeded;

        }
    }
}

