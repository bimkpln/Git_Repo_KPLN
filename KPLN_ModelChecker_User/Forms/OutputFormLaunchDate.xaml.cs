using Autodesk.Revit.DB;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.Common;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class OutputFormLaunchDate : Window
    {
        public OutputFormLaunchDate(Document doc, ExtensibleStorageEntity[] esEntitys)
        {
            Array.Sort(esEntitys, (x, y) => string.Compare(x.CheckName, y.CheckName));

            InitializeComponent();

            // Записываю данные по последним запускам
            Element piElem = doc.ProjectInformation;
            foreach (ExtensibleStorageEntity entity in esEntitys)
            {
                ResultMessage esMsgRun = entity.ESBuilderRun.GetResMessage_Element(piElem);
                entity.LastRunText = esMsgRun.Description;
                if (entity.LastRunText.Equals($"Данные отсутсвуют (не запускался)"))
                    entity.TextColor = Brushes.Red;
                else
                    entity.TextColor = Brushes.Green;
            }

            modulesListBox.ItemsSource = esEntitys;
            DataContext = this;
        }
    }
}
