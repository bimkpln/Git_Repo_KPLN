using Autodesk.Revit.UI;
using KPLN_Tools_AR.Forms.Models;
using System.Windows;

namespace KPLN_Tools_AR.Forms
{
    /// <summary>
    /// Логика взаимодействия для AR_TEPDesign_categorySelect.xaml
    /// </summary>
    public partial class AR_PyatnGraph_Main : Window
    {
        public AR_PyatnGraph_Main(UIApplication uiapp)
        {
            InitializeComponent();
            KPLN_Library_Forms.Services.WindowHandleSearch.MainWindowHandle.SetAsOwner(this);
            
            ARPG_VM = new AR_PyatnGraph_VM(uiapp);
            DataContext = ARPG_VM;
        }

        /// <summary>
        /// Ссылка на VM для окна
        /// </summary>
        public AR_PyatnGraph_VM ARPG_VM { get; private set; }
    }
}
