using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Forms.MVVM;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ListRenumberForm : Window
    {
        public ListRenumberForm(UIApplication uiapp, IEnumerable<ViewSheet> sortedSheets, ParameterSet tbs)
        {
            List<Parameter> mixedParamList = new List<Parameter>();
            foreach (Parameter curParam in tbs)
            {
                mixedParamList.Add(curParam);
            }

            InitializeComponent();

            var listRenVM = new ListRenumberVM(uiapp, sortedSheets, mixedParamList.OrderBy(p => p.Definition.Name));
            DataContext = listRenVM;
            if (listRenVM.CloseAction == null)
                listRenVM.CloseAction = new System.Action(ActionToClose);
        }

        private void ActionToClose()
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}
