using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.ExecutableCommand;
using KPLN_ModelChecker_Lib.Forms.Entities;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_ModelChecker_Lib.Forms
{
    /// <summary>
    /// Interaction logic for EmptyWorksetsForm.xaml
    /// </summary>
    public partial class EmptyWorksetsForm : Window
    {
        private readonly Document _doc;

        public ObservableCollection<WSEntity> EmptyWorksets { get; } = new ObservableCollection<WSEntity>();

        public string DocumentTitle { get; private set; }

        public EmptyWorksetsForm(Document doc)
        {
            InitializeComponent();

            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            DocumentTitle = _doc.Title;

            if (_doc.IsWorkshared)
            {
                Workset[] userWs = new FilteredWorksetCollector(_doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .ToWorksets()
                    .ToArray();

                foreach (Workset w in userWs)
                {
                    ElementWorksetFilter wfilter = new ElementWorksetFilter(w.Id);
                    FilteredElementCollector col = new FilteredElementCollector(doc).WherePasses(wfilter);
                    if (col.GetElementCount() == 0)
                        EmptyWorksets.Add(new WSEntity(w));
                }
            }

            DataContext = this;
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var toDelete = EmptyWorksets.Where(w => w.IsSelected).ToList();
            if (!toDelete.Any())
            {
                MessageBox.Show("Ничего не выбрано.", "Empty WS", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Удаляем
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new WSDeleter(_doc, toDelete.Select(ent => ent.RevitWS)));
        }
    }
}
