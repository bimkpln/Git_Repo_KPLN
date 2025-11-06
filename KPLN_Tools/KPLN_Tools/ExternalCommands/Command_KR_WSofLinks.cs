using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_KR_WSofLinks : IExternalCommand
    {
        internal const string PluginName = "Рабочие наборы связей";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //получение документа
            UIDocument uidoc = commandData.Application.ActiveUIDocument;   
            Document doc = uidoc.Document;

            //коллекция всех загруженных экземпляров связей
            var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                .Where(link => (doc.GetElement(link.GetTypeId()) as RevitLinkType).IsNestedLink == false);

            List<Document> linkedDocs = new List<Document>();

            //"имя для РН который хотим вкл/выкл";
            string targetWorksetName = null;
            //сценарий рабоыт плагина (открыть или закрыть введенный РН)
            bool WSOpenClose;

            var inputWindow = new KR_WSofLinks();
            bool? result = inputWindow.ShowDialog();

            //проверяем ввел ли пользователь имя РН, если нет, то отменяем работу плагина
            if (result == true)
            {
                targetWorksetName = inputWindow.WorksetName;
                WSOpenClose = inputWindow.WorksetOpenClose;
            }
            else
            {
                // Пользователь отменил ввод
                return Result.Cancelled;
            }

            if (string.IsNullOrWhiteSpace(targetWorksetName))   //проверяем, получил ли наш параметр какое-то значение ,если нет, то отменяем работу плагина
            {
                TaskDialog.Show("Ошибка", "Имя рабочего набора не задано.");
                return Result.Failed;
            }

            foreach (var linkInstance in linkInstances)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                // если ссылка на связь не корректная то пропускаем ее
                if (linkedDoc == null)
                    continue;


                // Получаем тип связи
                RevitLinkType linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                if (linkType == null) continue;

                ExternalFileReference extFileRef = linkType.GetExternalFileReference();
                if (extFileRef == null) continue;
                ModelPath absPath = extFileRef.GetAbsolutePath();

                // Формируем список Id рабочих наборов, которые нужно ОТКРЫТЬ (все, кроме целевого)
                var openedWorksetNames = new FilteredWorksetCollector(linkedDoc)
                                                                                .OfKind(WorksetKind.UserWorkset)
                                                                                .ToWorksets()
                                                                                .Where(ws => ws.IsOpen)
                                                                                .Select(ws => ws.Name)
                                                                                .ToHashSet();
                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(absPath);
                IList<WorksetId> worksetsToOpen = new List<WorksetId>();
                IList<WorksetPreview> worksetsTypeToOpen = new List<WorksetPreview>();

                foreach (var ws in worksets)
                {
                    //если выбран сценарий закрыть и соответственно галка не стоит
                    if (!WSOpenClose)
                    {
                        // Открываем только те, что были открыты и не равны целевому
                        if (ws.Name != targetWorksetName && openedWorksetNames.Contains(ws.Name))
                        {
                            worksetsToOpen.Add(ws.Id);
                            worksetsTypeToOpen.Add(ws);
                        }
                    }
                    else
                    {
                        // Открываем только те, что были открыты и равны целевому
                        if (ws.Name == targetWorksetName || openedWorksetNames.Contains(ws.Name))
                        {
                            worksetsToOpen.Add(ws.Id);
                            worksetsTypeToOpen.Add(ws);
                        }
                    }

                }

                // Конфигурация: по умолчанию все закрыты, открываем только нужные
                WorksetConfiguration wsConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                if (worksetsToOpen.Count > 0)
                    wsConfig.Open(worksetsToOpen);

                // Параметры загрузки связи с новой конфигурацией рабочих наборов
                RevitLinkOptions options = new RevitLinkOptions(true);
                options.SetWorksetConfiguration(wsConfig);

                // Выгружаем связь
                linkType.Unload(null);

                // Перезагружаем связь
                linkType.LoadFrom(absPath, wsConfig);
            }

            return Result.Succeeded;
        }
    }
}
