using Autodesk.Revit.UI;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Library_Forms.UIFactory
{
    /// <summary>
    // Инициализация загрузки окна выбора файлов Revit-Server
    /// </summary>
    public static class SelectFilesFromRevitServer
    {
        /// <summary>
        /// Ссылка на созданный РС для дальнейшей работы
        /// </summary>
        public static RevitServer CurrentRevitServer { get; private set; }

        public static ElementMultiPick CreateForm(int revitVersion)
        {
            #region Выбор РС
            ElementSinglePick selectedRevitServerMainDirForm = SelectRevitServerMainDir.CreateForm_SelectRSMainDir(revitVersion);
            bool? dialogResult = selectedRevitServerMainDirForm.ShowDialog();
            if (dialogResult == null || selectedRevitServerMainDirForm.Status != UIStatus.RunStatus.Run)
                return null;

            string selectedRSMainDirFullPath = selectedRevitServerMainDirForm.SelectedElement.Element as string;
            string selectedRSHostName = selectedRSMainDirFullPath.Split('\\')[0];
            string selectedRSMainDir = selectedRSMainDirFullPath.TrimStart(selectedRSHostName.ToCharArray());
            #endregion

            #region Выбор элементов с РС
            try
            {
                ObservableCollection<ElementEntity> projects = new ObservableCollection<ElementEntity>();

                CurrentRevitServer = new RevitServer(selectedRSHostName, revitVersion);
                FolderContents folderContents = CurrentRevitServer.GetFolderContents(selectedRSMainDir);
                List<Model> activeModelsFromMainDir = GetModelsFromMainDir(folderContents);

                IEnumerable<ElementEntity> activeEntitiesForForm = activeModelsFromMainDir.Select(e => new ElementEntity(e.Path));

                ElementMultiPick pickForm = new ElementMultiPick(activeEntitiesForForm.OrderBy(p => p.Name), "Выбери файлы");

                return pickForm;
            }
            catch (Exception ex)
            {
                TaskDialog td = new TaskDialog("KPLN: Ошибка")
                {
                    MainContent = $"При обращении к Revit-Server возникла ошибка:\n{ex.Message}",
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                };
                td.Show();

                return null;
            }
            #endregion

        }

        private static List<Model> GetModelsFromMainDir(FolderContents folderContents)
        {
            List<Model> models = new List<Model>();

            foreach (Model model in folderContents.Models)
            {
                if (model.LockState == LockState.Locked)
                    continue;

                models.Add(model);
            }
            foreach (Folder folder in folderContents.Folders)
            {
                if (folder.Name.ToLower().Contains("архив"))
                    continue;

                FolderContents recFolderContents = CurrentRevitServer.GetFolderContents(folder.Path);
                models.AddRange(GetModelsFromMainDir(recFolderContents));
            }

            return models;
        }
    }
}
