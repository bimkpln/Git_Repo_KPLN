#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion

#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
#endregion

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class FormBatchDelete : Form
    {
        private List<string> m_items = new List<string>();
        public List<string> Items
        {
            get { return m_items; }
            set { m_items = value ?? new List<string>(); }
        }

        private List<string> m_unusedItems = new List<string>();
        public List<string> UnusedItems
        {
            get { return m_unusedItems; }
            set { m_unusedItems = value ?? new List<string>(); }
        }

        private List<string> m_checkedItems = new List<string>();
        public List<string> CheckedItems
        {
            get { return m_checkedItems; }
        }

        public FormBatchDelete()
        {
            InitializeComponent();
        }

        private void FormBatchDelete_Load(object sender, EventArgs e)
        {
            UpdateList();
        }

        private void UpdateList()
        {
            listBox1.Items.Clear();

            List<string> source = checkBoxUnusedOnly.Checked
                ? m_unusedItems
                : m_items;

            listBox1.Items.AddRange(source.OrderBy(x => x).ToArray());
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            m_checkedItems.Clear();

            foreach (object item in listBox1.SelectedItems)
            {
                m_checkedItems.Add(item.ToString());
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBoxUnusedOnly_CheckedChanged(object sender, EventArgs e)
        {
            UpdateList();
        }
    }
}