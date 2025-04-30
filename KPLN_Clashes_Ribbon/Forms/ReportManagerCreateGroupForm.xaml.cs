using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_Bitrix24Worker;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportManagerCreateGroupForm.xaml
    /// </summary>
    public partial class ReportManagerCreateGroupForm : System.Windows.Window
    {
        public ReportManagerCreateGroupForm()
        {
            InitializeComponent();

            DataContext = CurrentReportGroup;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public ReportGroup CurrentReportGroup { get; private set; } = new ReportGroup();

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(ReportNameTBl);

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        private void LoadByParentBtn_Click(object sender, RoutedEventArgs args)
        {
            if (CurrentDBUser.SubDepartmentId != 8) { return; }

            TextInputForm textInputForm = new TextInputForm(this, "Введите ID базавой задачи Bitrix:");
            if ((bool)textInputForm.ShowDialog())
            {
                if (!int.TryParse(textInputForm.UserComment, out int bitrParentTaskId))
                {
                    System.Windows.MessageBox.Show(
                        $"Можно вводить ТОЛЬКО числа",
                        "Ошибка поиска задач в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                Task<Dictionary<int, string>> taskIdAndTitleDict = Task<Dictionary<int, string>>.Run(() =>
                {
                    return BitrixMessageSender
                        .GetAllSubTasks_IdAndTitle_ByParentId(bitrParentTaskId);
                });
                Dictionary<int, string> bitrTaskIdDicts = taskIdAndTitleDict.Result;
                if (bitrTaskIdDicts == null || !bitrTaskIdDicts.Any())
                {
                    System.Windows.MessageBox.Show(
                        $"Либо задачи с введенным ID не существует, ЛИБО указанный ID соответсвует задаче БЕЗ подзадач",
                        "Ошибка поиска задач в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                SetBitrixIdTasks(bitrTaskIdDicts);
            }

            DataContext = CurrentReportGroup;
        }

        private void SetBitrixIdTasks(Dictionary<int, string> bitrTaskIdDicts)
        {
            foreach (KeyValuePair<int, string> kvp in bitrTaskIdDicts)
            {
                string taskTitle = kvp.Value;
                if (taskTitle.Contains("для АР."))
                    CurrentReportGroup.BitrixTaskIdAR = kvp.Key;
                else if (taskTitle.Contains("для КР."))
                    CurrentReportGroup.BitrixTaskIdKR = kvp.Key;
                else if (taskTitle.Contains("для ОВ."))
                    CurrentReportGroup.BitrixTaskIdOV = kvp.Key;
                else if (taskTitle.Contains("для ИТП."))
                    CurrentReportGroup.BitrixTaskIdITP = kvp.Key;
                else if (taskTitle.Contains("для ВК."))
                    CurrentReportGroup.BitrixTaskIdVK = kvp.Key;
                else if (taskTitle.Contains("для АУПТ."))
                    CurrentReportGroup.BitrixTaskIdAUPT = kvp.Key;
                else if (taskTitle.Contains("для ЭОМ."))
                    CurrentReportGroup.BitrixTaskIdEOM = kvp.Key;
                else if (taskTitle.Contains("для СС."))
                    CurrentReportGroup.BitrixTaskIdSS = kvp.Key;
                else if (taskTitle.Contains("для СС_АВ."))
                    CurrentReportGroup.BitrixTaskIdAV = kvp.Key;
            }
        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentReportGroup.Name) || CurrentReportGroup.Name.Length < 5)
            {
                MessageBox.Show(
                    $"Имя группы не может быть пустым, или слишком коротким (меньше 5-ти символов)",
                    "Ошибка создания группы",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return;
            }
            
            this.DialogResult = true;
            Close();
        }
    }
}
