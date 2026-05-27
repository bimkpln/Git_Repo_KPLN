using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.ExecutableCommand;
using KPLN_ModelChecker_Lib.Forms.Entities;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ModelChecker_Lib.Forms
{
    /// <summary>
    /// Interaction logic for EmptyWorksetsForm.xaml
    /// </summary>
    public partial class WorksetsForm : Window
    {
        private readonly Document _doc;
        private int _lastCheckedIndex = -1;
        private readonly Workset[] _allWorksets;

        public WorksetsForm(Document doc, Workset[] wsColl)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            DocumentTitle = _doc.Title;

            _allWorksets = Util.GetDocWorksets(_doc) ?? new Workset[0];
            foreach (var ws in wsColl.OrderBy(ws => ws.Name))
            {
                WSColl.Add(new WSEntity(_doc, ws, _allWorksets));
            }

            InitializeComponent();
            DataContext = this;
        }

        public string DocumentTitle { get; private set; }

        public ObservableCollection<WSEntity> WSColl { get; } = new ObservableCollection<WSEntity>();

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var toDelete = WSColl.Where(w => w.IsSelected).ToList();
            if (!toDelete.Any())
                return;

            // Подготовка словаря
            var replacementMap = toDelete.ToDictionary(ent => ent.RevitWS, ent => ent.SelectedReplacementWS);
            

            // Всем РН назначена замена?
            if (toDelete.Any(ent => (ent.SelectedReplacementWS == null && ent.RevitWSElemsCount > 0)))
            {
                MessageBox.Show(
                    "Укажите РН на замену для всех выбранных рабочих наборов.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }


            // Нельзя удалить вообще все РН
            if (replacementMap.Count >= _allWorksets.Count())
            {
                MessageBox.Show("Нельзя удалить все пользовательские рабочие наборы.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }


            // Нет ли среди удаляемых РН те, которые выбраны на замену
            var toDeleteIds = new HashSet<WorksetId>(toDelete.Select(ent => ent.RevitWS.Id));

            var conflicts = toDelete
                .Where(ent => ent.SelectedReplacementWS != null && toDeleteIds.Contains(ent.SelectedReplacementWS.Id))
                .Select(ent => $"  • «{ent.RevitWS.Name}» — замена «{ent.SelectedReplacementWS.Name}», который помечен на удаление")
                .ToList();

            if (conflicts.Any())
            {
                MessageBox.Show(
                    $"Конфликт в выборе замены:\n\n{string.Join("\n", conflicts)}\n\nУкажите другой РН на замену.",
                    "Конфликт",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }


            // Удаляю
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new WSDeleter(_doc, replacementMap));

            this.Close();
        }

        private void WorksetsGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (e.Column.Header?.ToString() == "РН на замену")
            {
                if (e.Row.Item is WSEntity entity && entity.RevitWSElemsCount == 0)
                    e.Cancel = true;
            }
        }

        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is DataGridCell cell) || cell.Column.Header?.ToString() != "Удалить?")
                return;
            if (!(cell.DataContext is WSEntity entity))
                return;

            bool newState = !entity.IsSelected;
            int currentIndex = WSColl.IndexOf(entity);

            if (Keyboard.Modifiers == ModifierKeys.Shift && _lastCheckedIndex >= 0)
            {
                int from = Math.Min(_lastCheckedIndex, currentIndex);
                int to = Math.Max(_lastCheckedIndex, currentIndex);
                for (int i = from; i <= to; i++)
                    WSColl[i].IsSelected = newState;
            }
            else
            {
                entity.IsSelected = newState;
            }

            _lastCheckedIndex = currentIndex;
            e.Handled = true;
        }
    }
}