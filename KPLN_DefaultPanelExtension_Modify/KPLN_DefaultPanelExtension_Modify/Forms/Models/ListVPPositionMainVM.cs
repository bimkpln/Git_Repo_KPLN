using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_DefaultPanelExtension_Modify.ExecutableCommands;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class ListVPPositionMainVM
    {
        private readonly Window _owner;
        private readonly string _cofigName = "ViewportPositionConfig";
        private readonly int _configItemsHash;

        public ListVPPositionMainVM(Window owner, Element[] selTrueElems)
        {
            _owner = owner;
            SelectedViewElems = selTrueElems;
            CheckANDSetProcessedViewport();

            SelectCmd = new RelayCommand<ListVPPositionCreateM>(Select);
            CreateCmd = new RelayCommand<object>(_ => Create());
            HelpCmd = new RelayCommand<object>(_ => Help());
            DeleteCmd = new RelayCommand<ListVPPositionCreateM>(Delete);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);

            // Чтение конфигурации последнего запуска
            object lastRunConfigObj = ConfigService.ReadConfigFile<ListVPPositionCreateM[]>(ModuleData.RevitVersion, ProcessedViewElems.Document, ConfigType.Shared, _cofigName);
            if (lastRunConfigObj != null && lastRunConfigObj is ListVPPositionCreateM[] coll)
            {
                var orderedColl = coll.OrderBy(cm => cm.ConfigName);
                foreach (ListVPPositionCreateM item in orderedColl)
                {
                    ListVPPositionCreateItems.Add(item);
                }

                _configItemsHash = GetItemsHash();
            }
        }

        /// <summary>
        /// Выбранные виды/спеки
        /// </summary>
        public Element[] SelectedViewElems { get; set; }

        /// <summary>
        /// Обрабатываемый вид/спека
        /// </summary>
        public Element ProcessedViewElems { get; private set; }

        public ICommand SelectCmd { get; }

        public ICommand CreateCmd { get; }

        public ICommand HelpCmd { get; }

        public ICommand DeleteCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public ObservableCollection<ListVPPositionCreateM> ListVPPositionCreateItems { get; set; } = new ObservableCollection<ListVPPositionCreateM>();

        public void SaveConfig()
        {
            int currentHash = GetItemsHash();
            if (currentHash != _configItemsHash)
                ConfigService.SaveConfig<ListVPPositionCreateM>(ModuleData.RevitVersion, ProcessedViewElems.Document, ConfigType.Shared, ListVPPositionCreateItems, _cofigName);
        }

        /// <summary>
        /// Проверка и установка значения обрабатываемого вида
        /// </summary>
        private bool CheckANDSetProcessedViewport()
        {
            if (SelectedViewElems.Length == 0)
            {
                MessageBox.Show(
                    "В выборку не попали видовые экраны. Работа остановлена",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return false;
            }

            if (SelectedViewElems.Count() != 1)
            {
                MessageBox.Show(
                    $"Работать с конфигурациями можно только на основе одного вида, а сейчас выделено {SelectedViewElems.Count()} шт.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return false;
            }


            ProcessedViewElems = SelectedViewElems[0];
            return true;
        }

        private void Select(ListVPPositionCreateM item)
        {
            if(CheckANDSetProcessedViewport())
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdListVPPositionSet(ProcessedViewElems, item));
        }

        private void Create()
        {
            if (!CheckANDSetProcessedViewport())
                return;


            ListVPPositionCreateFrom createFrom = new ListVPPositionCreateFrom(_owner, ProcessedViewElems);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(createFrom);

            if ((bool)createFrom.ShowDialog())
                ListVPPositionCreateItems.Add(createFrom.CurentListVPPositionCreateVM.ListVPPositionItem);
        }

        private void Help() =>
            Process.Start(new ProcessStartInfo(ExcCmdListVPPositionStart.HelpUrl) { UseShellExecute = true });

        private void Delete(ListVPPositionCreateM item)
        {
            var td = MessageBox.Show(_owner, $"Сейчас из будет удалена конфигурация \"{item.ConfigName}\"", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (td == MessageBoxResult.Yes)
            {
                ListVPPositionCreateItems.Remove(item);
                SaveConfig();
            }
        }

        private void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        private int GetItemsHash()
        {
            int result = 0;

            if (ListVPPositionCreateItems.Count > 0)
            {
                foreach (var item in ListVPPositionCreateItems)
                {
                    result ^= item.GetHashCode();
                }
            }

            return result;
        }
    }
}
