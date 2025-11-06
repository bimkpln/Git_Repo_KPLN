using Autodesk.Revit.UI;
using KPLN_Tools.Forms.Models;
using System.Windows;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для AR_TEPDesign_categorySelect.xaml
    /// </summary>
    public partial class AR_PyatnGraph_Main : Window
    {
        public AR_PyatnGraph_Main(UIApplication uiapp)
        {
            InitializeComponent();
            
            ARPG_VM = new AR_PyatnGraph_VM(uiapp);
            DataContext = ARPG_VM;
        }

        /// <summary>
        /// Ссылка на VM для окна
        /// </summary>
        public AR_PyatnGraph_VM ARPG_VM { get; private set; }
    }
}
