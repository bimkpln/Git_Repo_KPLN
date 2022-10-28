using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using KPLN_ModelChecker_Lib.Common;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;
using System.Linq;
using KPLN_ModelChecker_Lib.Common.ErrorTypes;

namespace KPLN_ModelChecker_User.Common
{
    internal sealed class WPFReport
    {
        private readonly IList<ElementEntity> _errorList;

        public WPFReport(IList<ElementEntity> errorList)
        {
            _errorList = errorList;
        }

        /// <summary>
        /// Метод для вывода результатов пользователю
        /// </summary>
        public void ShowResult()
        {
            if (_errorList.Count == 0)
            {
                TaskDialog.Show("Результат", "Проблемы не обнаружены :)", TaskDialogCommonButtons.Ok);
            }
            else
            {
                int counter = 1;
                List<ElementEntity> sortedOutputCollection = _errorList.OrderBy(e => e.ErrorStatus.Id).ToList();
                foreach (ElementEntity e in sortedOutputCollection)
                {
                    e.Title = string.Format("{0}# {1}", (counter++).ToString(), e.Title);
                }

                List<ElementEntity> wpfFiltration = new List<ElementEntity>();
                wpfFiltration.Add(new ElementEntity(null, "<Все>", null, null, null ));
                
                if (sortedOutputCollection.Count != 0)
                {
                    OutputForm form = new OutputForm(sortedOutputCollection, wpfFiltration);
                    form.Show();
                }
            }
        }
    }
}
