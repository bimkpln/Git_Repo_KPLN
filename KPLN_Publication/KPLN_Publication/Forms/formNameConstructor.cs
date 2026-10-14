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

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System.Windows.Interop;

namespace KPLN_Publication
{
    public partial class formNameConstructor : System.Windows.Forms.Form
    {
        public string nameConstructor;
        private readonly IEnumerable<MainEntity> _entities;

        public formNameConstructor(string NameConstructor, IEnumerable<MainEntity> entities)
        {
            _entities = entities.ToList();
            InitializeComponent();
            this.AcceptButton = buttonOk;
            this.CancelButton = buttonCancel;
            textBoxNameConstructor.Text = NameConstructor;
        }

        private void buttonParameter_Click(object sender, EventArgs e)
        {
            int selectionStart = textBoxNameConstructor.SelectionStart;
            int selectionLength = textBoxNameConstructor.SelectionLength;
            Dictionary<string, HashSet<string>> parameters = new Dictionary<string, HashSet<string>>();
            foreach (MainEntity entity in _entities)
            {
                foreach (Element source in MainEntity.GetNameParameterSources(entity.MainView))
                {
                    string origin = source is ElementType ? "Тип" : "Экземпляр";
                    origin += ": " + source.Name + " (" + source.Document.Title + ")";
                    foreach (Parameter parameter in source.Parameters)
                    {
                        if (parameter.StorageType == StorageType.None)
                            continue;
                        string name = parameter.Definition.Name;
                        if (!parameters.ContainsKey(name))
                            parameters.Add(name, new HashSet<string>());
                        parameters[name].Add(origin);
                    }
                }
            }

            // Библиотечное окно с поиском. Берём объединение параметров всех источников,
            // а не пересечение листа с информацией о проекте.
            IEnumerable<ElementEntity> choices = parameters.OrderBy(pair => pair.Key)
                .Select(pair => new ElementEntity(pair.Key, string.Join("\n", pair.Value)));
            ElementSinglePick picker = new ElementSinglePick(null, choices, "Параметр имени файла");
            new WindowInteropHelper(picker) { Owner = Handle };
            if (picker.ShowDialog() != true || picker.SelectedElement == null)
                return;

            textBoxNameConstructor.Select(selectionStart, selectionLength);
            textBoxNameConstructor.SelectedText = "<" + picker.SelectedElement.Name + ">";
            textBoxNameConstructor.Focus();
        }


        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            nameConstructor = textBoxNameConstructor.Text;
            this.Close();
        }
    }
}
