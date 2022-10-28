using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ModelChecker_Coordinator.Forms
{
    public class DbTypeGraph : INotifyPropertyChanged
    {
        public DbTypeGraph(ChartValues<DateTimePoint> data, string name)
        {
            Data = data;
            Name = name;
            Line = new LineSeries
            {
                Title = Name,
                Values = Data,
                PointGeometrySize = 5
            };
        }
        public LineSeries Line { get; set; }
        public ChartValues<DateTimePoint> Data { get; set; }
        public string Name { get; }
        private bool _isChecked = true;
        public bool IsChecked
        {
            get
            {
                return _isChecked;
            }
            set
            {
                if (value == _isChecked)
                {
                    return;
                }
                _isChecked = value;
                NotifyPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
