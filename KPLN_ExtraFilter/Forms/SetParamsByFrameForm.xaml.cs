using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SetParamsByFrameForm : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly Element[] _elemsToSet;
        private string _runButtonName;
        private string _runButtonTooltip;

        /// <summary>
        /// Конструктор с открытием пустого окна
        /// </summary>
        public SetParamsByFrameForm(IEnumerable<Element> elemsToSet, IEnumerable<ParamEntity> paramsEntities)
        {
            _elemsToSet = elemsToSet.ToArray();
            AllParamEntities = paramsEntities.OrderBy(ent => ent.CurrentParamName).ToArray();

            InitializeComponent();
            RunButtonContext();

            DataContext = this;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Конструктор с открытием окна с преднастройкой
        /// </summary>
        public SetParamsByFrameForm(
            IEnumerable<Element> elemsToSet,
            IEnumerable<ParamEntity> paramsEntities,
            IEnumerable<MainItem> userMainItem) : this(elemsToSet, paramsEntities)
        {
            foreach (MainItem mi in userMainItem)
            {
                MainItems.Add(mi);
            }

            RunButtonContext();
        }

        public string RunButtonName
        {
            get => _runButtonName;
            private set
            {
                if (_runButtonName != value)
                {
                    _runButtonName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string RunButtonTooltip
        {
            get => _runButtonTooltip;
            private set
            {
                if (_runButtonTooltip != value)
                {
                    _runButtonTooltip = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коллекция ВСЕХ сущностей параметров, которые есть у эл-в
        /// </summary>
        public ParamEntity[] AllParamEntities { get; private set; }

        public ObservableCollection<MainItem> MainItems { get; private set; } = new ObservableCollection<MainItem>();

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        /// <summary>
        /// Установить данные по кнопке
        /// </summary>
        private void RunButtonContext()
        {
            if (MainItems.Count > 0)
            {
                RunButtonName = "Заполнить!";
                RunButtonTooltip = "Заполнить указанные параметры указанными значениями для выделенных элементов";
            }
            else
            {
                RunButtonName = "Выделить!";
                RunButtonTooltip = "Выделить все вложенности у выбранных элементов";
            }
        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue
                .Enqueue(new SetParamsByFrameExcCommandStart(_elemsToSet, MainItems));

            Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            MainItem defaultMI = new MainItem(AllParamEntities.FirstOrDefault());
            MainItems.Add(defaultMI);

            RunButtonContext();
        }

        private void ClearBtn_Click(object sender, RoutedEventArgs e)
        {
            MainItems.Clear();

            RunButtonContext();
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if (!((MenuItem)e.Source is MenuItem menuItem))
                return;

            if (!(menuItem.DataContext is MainItem entity))
                return;

            UserDialog ud = new UserDialog("ВНИМАНИЕ",
                $"Сейчас будут удален параметр \"{entity.UserSelectedParamEntity.CurrentParamName}\". Продолжить?");
            ud.ShowDialog();

            if (ud.IsRun)
                MainItems.Remove(entity);

            RunButtonContext();
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
