using Autodesk.Revit.DB;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.Common;
using System;
using System.Linq;
using System.Text;
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
                string resultMsg = string.Empty;
                
                int maxLines = 5;
                ResultMessage esMsgRun = entity.ESBuilderRun.GetResMessage_Element(piElem);
                string esMsgDiscr = esMsgRun.Description;
                
                string[] splitedMsgDescr = esMsgDiscr.Split('\n');
                int lineCount = splitedMsgDescr.Count();
                if (lineCount > maxLines)
                {
                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine($"<<<Всего больше {maxLines} запусков. Остальное можешь посмотреть через RevitLookup>>>");
                    for (int i = maxLines; i > 0; i--)
                    {
                        stringBuilder.AppendLine(splitedMsgDescr[lineCount - i]);
                    }
                    resultMsg = stringBuilder.ToString();
                }
                else
                    resultMsg = esMsgDiscr;


                entity.LastRunText = resultMsg;
                if (entity.LastRunText.Equals($"Данные отсутствуют (не запускался)"))
                    entity.TextColor = Brushes.Red;
                else
                    entity.TextColor = Brushes.Green;
            }

            ModuleLogsListBox.ItemsSource = esEntitys;
            DataContext = this;
        }

        private void ModulesListBox_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            // Передаём событие ScrollViewer'у
            if (ModuleLogsStackScroll != null)
            {
                // Если крутим вниз, увеличиваем вертикальное смещение
                if (e.Delta < 0)
                {
                    ModuleLogsStackScroll.ScrollToVerticalOffset(ModuleLogsStackScroll.VerticalOffset + 20); // Примерный шаг прокрутки
                }
                // Если крутим вверх, уменьшаем вертикальное смещение
                else if (e.Delta > 0)
                {
                    ModuleLogsStackScroll.ScrollToVerticalOffset(ModuleLogsStackScroll.VerticalOffset - 20);
                }

                // Указываем, что событие обработано
                e.Handled = true;
            }
        }
    }
}
