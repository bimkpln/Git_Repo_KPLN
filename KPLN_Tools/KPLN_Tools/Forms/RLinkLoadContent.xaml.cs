using KPLN_Library_Forms.UI;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.Forms.Models.Core;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class RLinkLoadContent : UserControl, IRLinkUserControl
    {
        public RLinkLoadContent()
        {
            LinkChangeEntityColl = new ObservableCollection<LinkManagerEntity>();

            InitializeComponent();
            DataContext = this;
        }

        public ObservableCollection<LinkManagerEntity> LinkChangeEntityColl { get; set; }

        public void AddNewItem(LinkManagerEntity entity)
        {
            if (entity is LinkManagerLoadEntity loadEnt)
            {
                if (!LinkChangeEntityColl.Any(ent => ent.LinkPath == loadEnt.LinkPath))
                    LinkChangeEntityColl.Add(loadEnt);
                else
                {
                    CustomMessageBox cmb = new CustomMessageBox(
                    "Предупреждение",
                        $"Файл \"{loadEnt.LinkName}\" по пути \"{loadEnt.LinkPath}\" уже присутсвует в конфигурации. Дублирование запрещено");
                    cmb.ShowDialog();
                }
            }
        }

        public void RemoveItem(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.DataContext is LinkManagerLoadEntity linkLoad)
            {
                UserDialog ud = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалена загрузка файла \"{linkLoad.LinkName}\". Продолжить?");
                ud.ShowDialog();

                if (ud.IsRun)
                    LinkChangeEntityColl.Remove(linkLoad);
            }
        }

        private void CopyFromFirst_Click(object sender, RoutedEventArgs e)
        {
            LinkManagerLoadEntity linkLoad = LinkChangeEntityColl.Cast<LinkManagerLoadEntity>().FirstOrDefault();
            if (linkLoad == null)
                return;

            CoordinateType firstCoordType = linkLoad.LinkCoordinateType;
            bool firstCreateWS = linkLoad.CreateWorksetForLinkInst;
            string firstClosedWS = linkLoad.WorksetToCloseNamesStartWith;

            foreach (LinkManagerLoadEntity ent in LinkChangeEntityColl.Cast<LinkManagerLoadEntity>()) 
            {
                ent.LinkCoordinateType = firstCoordType;
                ent.CreateWorksetForLinkInst = firstCreateWS;
                ent.WorksetToCloseNamesStartWith = firstClosedWS;
            }
        }
    }
}
