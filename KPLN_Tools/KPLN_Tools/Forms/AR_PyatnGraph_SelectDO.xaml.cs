using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.AR_PyatnGraph;
using KPLN_Tools.Forms.Models;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для AR_TEPDesign_categorySelect.xaml
    /// </summary>
    public partial class AR_PyatnGraph_SelectDO : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ARPG_DesOptEntity _selARPGDesignOpt;
        private bool _canRun;

        public AR_PyatnGraph_SelectDO(DesignOption[] designOptions)
        {
            DocARPGDesignOpts = new ARPG_DesOptEntity[designOptions.Length + 1];
            DocARPGDesignOpts[0] = new ARPG_DesOptEntity() 
            { 
                ARPG_DesignOptionName = "Главная модель",
                ARPG_DesignOptionId = -1,
            };
            
            for (int i = 1; i < designOptions.Length + 1; i++)
            {
                DocARPGDesignOpts[i] = new ARPG_DesOptEntity()
                {
                    ARPG_DesignOptionName = designOptions[i - 1].Name,
                    ARPG_DesignOptionId = designOptions[i - 1].Id.IntegerValue,
                };

            }

            InitializeComponent();

            DataContext = this;
        }

        public ARPG_DesOptEntity[] DocARPGDesignOpts { get; private set; }

        public ARPG_DesOptEntity SelARPGDesignOpt
        {
            get => _selARPGDesignOpt;
            set
            {
                _selARPGDesignOpt = value;
                NotifyPropertyChanged();

                if (_selARPGDesignOpt != null)
                    CanRun = true;
                else
                    CanRun = false;
            }
        }

        /// <summary>
        /// Метка возможности запуска
        /// </summary>
        public bool CanRun
        {
            get => _canRun;
            set
            {
                _canRun = value;
                NotifyPropertyChanged();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            this.Close();
        }
    }
}
