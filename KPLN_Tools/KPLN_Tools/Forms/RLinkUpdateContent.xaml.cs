using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.Forms.Models.Core;
using Microsoft.Win32;
using RevitServerAPILib;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class RLinkUpdateContent : UserControl, IRLinkUserControl
    {
        private readonly int _revitVersion;

        public RLinkUpdateContent(int revitVersion)
        {
            LinkChangeEntityColl = new ObservableCollection<LinkManagerEntity>();
            _revitVersion = revitVersion;

            InitializeComponent();
            DataContext = this;
        }

        public ObservableCollection<LinkManagerEntity> LinkChangeEntityColl { get; set; }

        public void AddNewItem(LinkManagerEntity entity) => LinkChangeEntityColl.Add(entity);

        public void RemoveItem(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.DataContext is LinkManagerUpdateEntity linkChange)
            {
                UserDialog ud = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалена замена файла \"{linkChange.LinkName}\". Продолжить?");
                ud.ShowDialog();

                if (ud.IsRun)
                    LinkChangeEntityColl.Remove(linkChange);
            }
        }

        private void MarkAsFinal(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.DataContext is LinkManagerUpdateEntity linkChange)
            {
                if (linkChange.CurrentEntStatus == EntityStatus.Error)
                    return;
                else if (linkChange.CurrentEntStatus == EntityStatus.MarkedAsFinal)
                    linkChange.CurrentEntStatus = EntityStatus.Ok;
                else
                    linkChange.CurrentEntStatus = EntityStatus.MarkedAsFinal;
            }
        }

        private void ServerPathSelect_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog browserDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = RLinkManagerForm.InitialDirectoryForOpenFileDialog,
            };

            if (browserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                RLinkManagerForm.InitialDirectoryForOpenFileDialog = browserDialog.SelectedPath;

                UpdateCollByModelsFromSelectedPath(RLinkManagerForm.InitialDirectoryForOpenFileDialog);
            }
        }

        private void RevitServerPathSelect_Click(object sender, RoutedEventArgs e)
        {
            ElementSinglePick selectedRevitServerMainDirForm = SelectRevitServerMainDir.CreateForm_SelectRSMainDir(_revitVersion);
            bool? dialogResult = selectedRevitServerMainDirForm.ShowDialog();
            if (dialogResult == null || selectedRevitServerMainDirForm.Status != UIStatus.RunStatus.Run)
                return;

            string selectedRSMainDirFullPath = selectedRevitServerMainDirForm.SelectedElement.Element as string;
            string selectedRSHostName = selectedRSMainDirFullPath.Split('\\')[0];
            string selectedRSMainDir = selectedRSMainDirFullPath.TrimStart(selectedRSHostName.ToCharArray());

            RevitServer revitServer = new RevitServer(selectedRSHostName, _revitVersion);

            IList<Folder> rsFolders = revitServer.GetFolderContents(selectedRSMainDir, 0).Folders;
            List<ElementEntity> activeEntitiesForForm = new List<ElementEntity>(
                rsFolders
                    .Where(f => f.LockState != LockState.Locked)
                    .Select(f => new ElementEntity(f.Path))
                    .ToArray());

            ElementSinglePick pickForm = new ElementSinglePick(activeEntitiesForForm.OrderBy(p => p.Name), "Выбери папку Revit-Server");
            bool? pickFormResult = pickForm.ShowDialog();
            if (pickFormResult == null || pickForm.Status != UIStatus.RunStatus.Run)
                return;

            string dirToRSModels = $"\\\\{selectedRSHostName}{pickForm.SelectedElement.Name}";

            UpdateCollByModelsFromSelectedPath(dirToRSModels);
        }

        private void UpdateServLink_Click(object sender, RoutedEventArgs e)
        {
            Button lmeBtn = sender as Button;
            if (lmeBtn.DataContext is LinkManagerUpdateEntity currentLME)
            {
                if (currentLME.CurrentEntStatus == EntityStatus.MarkedAsFinal)
                {
                    UserDialog ud = new UserDialog(
                        "ВНИМАНИЕ",
                        $"Данная сущность помечена как итоговая. Операция отменена",
                        "Сними галку, если нужно перезаписать значение");
                    ud.ShowDialog();

                    return;
                }

                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Multiselect = false,
                    Filter = "Revit Files (*.rvt)|*.rvt",
                    InitialDirectory = RLinkManagerForm.InitialDirectoryForOpenFileDialog,
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    RLinkManagerForm.InitialDirectoryForOpenFileDialog = openFileDialog.FileName;

                    string filePath = openFileDialog.FileName;
                    string fileName = Path.GetFileName(filePath);
                    LinkManagerUpdateEntity newEntity = new LinkManagerUpdateEntity(currentLME.LinkName, currentLME.LinkPath, fileName, filePath);

                    ReloadWithNewData(currentLME, newEntity);
                }
            }
        }

        private void UpdateRevitServerLink_Click(object sender, RoutedEventArgs e)
        {
            Button lmeBtn = sender as Button;
            if (lmeBtn.DataContext is LinkManagerUpdateEntity currentLME)
            {
                if (currentLME.CurrentEntStatus == EntityStatus.MarkedAsFinal)
                {
                    UserDialog ud = new UserDialog(
                        "ВНИМАНИЕ",
                        $"Данная сущность помечена как итоговая. Операция отменена",
                        "Сними галку, если нужно перезаписать значение");
                    ud.ShowDialog();

                    return;
                }

                // Тут нужно заменить на одиночный выбор, но это нужно библиотеку править. Пока оставляю так
                ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
                if (rsFilesPickForm == null)
                    return;

                bool? dialogResult = rsFilesPickForm.ShowDialog();
                if (dialogResult == null || rsFilesPickForm.Status != UIStatus.RunStatus.Run)
                    return;

                string fileName = rsFilesPickForm.SelectedElements.FirstOrDefault().Name;
                string filePath = $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{rsFilesPickForm.SelectedElements.FirstOrDefault().Name}";
                LinkManagerUpdateEntity newEntity = new LinkManagerUpdateEntity(currentLME.LinkName, currentLME.LinkPath, fileName, filePath);

                ReloadWithNewData(currentLME, newEntity);
            }
        }

        private void UpdateCollByModelsFromSelectedPath(string selectedPath)
        {
            // Готовим коллекцию для замены
            Dictionary<int, LinkManagerUpdateEntity> resulDict = new Dictionary<int, LinkManagerUpdateEntity>();
            foreach (LinkManagerUpdateEntity lmEntity in LinkChangeEntityColl)
            {
                if (lmEntity.CurrentEntStatus == EntityStatus.MarkedAsFinal)
                    continue;

                LinkManagerUpdateEntity updatedEntity = LoadRLI_Service.GetSimilarByPath(lmEntity, selectedPath, _revitVersion);
                if (updatedEntity != null)
                    resulDict.Add(LinkChangeEntityColl.IndexOf(lmEntity), updatedEntity);
            }

            // Меняем коллекцию в окне
            foreach (KeyValuePair<int, LinkManagerUpdateEntity> kvp in resulDict)
            {
                int lmIndexToUpdate = kvp.Key;
                LinkManagerUpdateEntity lmEntityToUpdate = LinkChangeEntityColl[lmIndexToUpdate] as LinkManagerUpdateEntity;
                LinkManagerUpdateEntity updatedLMEntity = kvp.Value;

                ReloadWithNewData(lmEntityToUpdate, updatedLMEntity);
            }
        }

        private void ReloadWithNewData(LinkManagerUpdateEntity sourceLME, LinkManagerUpdateEntity targetLME)
        {
            int updateIndex = LinkChangeEntityColl.IndexOf(sourceLME);
            LinkChangeEntityColl.Remove(sourceLME);
            LinkChangeEntityColl.Insert(updateIndex, targetLME);

            DataContext = this;
        }
    }
}
