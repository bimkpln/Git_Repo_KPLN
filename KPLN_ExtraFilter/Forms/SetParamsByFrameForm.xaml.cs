using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SetParamsByFrameForm : Window
    {
        private readonly Element[] _elemsToSet;
        
        /// <summary>
        /// Конструктор с открытием пустого окна
        /// </summary>
        public SetParamsByFrameForm(IEnumerable<Element> elemsToSet, IEnumerable<ParamEntity> paramsEntities)
        {
            _elemsToSet = elemsToSet.ToArray();
            AllParamEntities = paramsEntities.OrderBy(ent => ent.CurrentParamName).ToArray();
            
            InitializeComponent();

            DataContext = this;
        }

        /// <summary>
        /// Конструктор с открытием окна с преднастройкой
        /// </summary>
        public SetParamsByFrameForm(
            IEnumerable<Element> elemsToSet, 
            IEnumerable<ParamEntity> paramsEntities, 
            IEnumerable<MainItem> userMainItem) : this(elemsToSet, paramsEntities)
        {
            foreach(MainItem mi in userMainItem)
            {
                MainItems.Add(mi);
            }
        }

        /// <summary>
        /// Коллекция ВСЕХ сущностей параметров, которые есть у эл-в
        /// </summary>
        public ParamEntity[] AllParamEntities { get; private set; }

        public ObservableCollection<MainItem> MainItems { get; private set; } = new ObservableCollection<MainItem>();

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
        }

        private void ClearBtn_Click(object sender, RoutedEventArgs e) => MainItems.Clear();

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
        }
    }
}
