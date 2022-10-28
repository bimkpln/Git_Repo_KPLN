using KPLN_ModelChecker_Coordinator.DB;
using LiveCharts;
using LiveCharts.Defaults;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_ModelChecker_Coordinator.Forms
{
    /// <summary>
    /// Логика взаимодействия для OutputDB.xaml
    /// </summary>
    public partial class OutputDB : Window
    {
        public ObservableCollection<DbTypeGraph> Data = new ObservableCollection<DbTypeGraph>();
        private ChartValues<DateTimePoint> GetData(List<DbRowData> data, string name)
        {
            var values = new ChartValues<DateTimePoint>();
            foreach (DbRowData row in data)
            {
                bool found = false;
                foreach (DbError err in row.Errors)
                {
                    if (err.Name == name)
                    {
                        found = true;
                        values.Add(new DateTimePoint(row.DateTime, err.Count));
                    }
                }
                if (!found)
                {
                    values.Add(new DateTimePoint(row.DateTime, double.NaN));
                }
            }
            if (values.Count == 1)
            {
                values.Add(new DateTimePoint(DateTime.Now, values[0].Value));
            }
            return values;
        }
        public string ProjectName { get; set; }
        public LiveCharts.SeriesCollection SeriesCollection { get; set; }
        public Func<double, string> XFormatter { get; set; }
        public Func<double, string> YFormatter { get; set; }
        public void Update()
        {
            foreach (DbTypeGraph type in Data)
            {
                if (type.IsChecked)
                {
                    if (!SeriesCollection.Contains(type.Line))
                    {
                        SeriesCollection.Add(type.Line);
                    }
                }
                else
                {
                    if (SeriesCollection.Contains(type.Line))
                    {
                        SeriesCollection.Remove(type.Line);
                    }
                }
            }
        }
        public static List<KPLN_Library_DataBase.Collections.DbDocument> _Documents { get; set; } = null;
        public OutputDB(string projectName, List<DbRowData> data, List<KPLN_Library_DataBase.Collections.DbDocument> documents)
        {
            Title = string.Format("KPLN: {0}", projectName);
            _Documents = documents;
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            ProjectName = projectName;
            InitializeComponent();
            HashSet<string> uniq = new HashSet<string>();
            foreach (DbRowData row in data)
            {
                foreach (DbError err in row.Errors)
                {
                    uniq.Add(err.Name);
                }
            }
            SeriesCollection = new SeriesCollection();
            foreach (string type in uniq)
            {
                Data.Add(new DbTypeGraph(GetData(data, type), type));
            }
            Update();
            icTypes.ItemsSource = Data;
            XFormatter = val => new DateTime((long)val).ToString("dd MMM");
            YFormatter = val => Math.Round(val, 0).ToString();
            DataContext = this;
            Closed += OnClosed;
        }
        private void OnClosed(object sender, EventArgs args)
        {
            if (_Documents != null)
            {
                Picker dp = new Picker(_Documents);
                dp.ShowDialog();
                if (Picker.PickedDocument != null)
                {
                    ProjectName = string.Format("{0}: {1}", Picker.PickedDocument.Project.Name, Picker.PickedDocument.Name);
                    KPLN_Library_DataBase.Collections.DbDocument pickedDocument = Picker.PickedDocument;
                    OutputDB form = new OutputDB(ProjectName, DbController.GetRows(pickedDocument.Id.ToString()), _Documents);
                    form.Show();
                }
            }
        }
        private void OnChecked(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void OnUnchecked(object sender, RoutedEventArgs e)
        {
            Update();
        }
    }

}
