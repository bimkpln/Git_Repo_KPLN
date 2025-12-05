using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KPLN_ViewsAndLists_Ribbon
{
    public partial class FormSelectCategories : System.Windows.Forms.Form
    {
        public List<ElementId> checkedCategoriesIds;

        public FormSelectCategories(List<ElementId> categoriesIds)
        {
            InitializeComponent();

            foreach (ElementId catId in categoriesIds)
            {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                BuiltInCategory bic = (BuiltInCategory)catId.IntegerValue;
#else
                BuiltInCategory bic = (BuiltInCategory)catId.Value;
#endif
                checkedListBox1.Items.Add(bic, CheckState.Checked);
            }

        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            checkedCategoriesIds = new List<ElementId>();

            foreach (var checkedItem in checkedListBox1.CheckedItems)
            {
                BuiltInCategory cat =(BuiltInCategory)checkedItem;

                checkedCategoriesIds.Add(new ElementId(cat));
            }

            this.DialogResult = DialogResult.OK;
        }
    }
}
