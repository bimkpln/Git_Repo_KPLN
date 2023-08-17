using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_Loader.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ProgressBar : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Прогресс бар по процессу
        /// </summary>
        
        public ProgressBar()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Маркер завершения подготовки окружения
        /// </summary>
        public bool EnvironmentDone { get; set; }

        /// <summary>
        /// Маркер завершения подключения к БД
        /// </summary>
        public bool UserConnectionDone { get; set; }

        /// <summary>
        /// Маркер завершения активации модулей
        /// </summary>
        public bool ModulesActivationDone { get; set; }

        /// <summary>
        /// Текстовое описание
        /// </summary>
        public string Instruction { get; set; }

    }
}
