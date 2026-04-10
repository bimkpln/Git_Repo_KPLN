using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        public ObservableCollection<LinkWorksetsItem> LinksWorksetsSP { get; set; } = new ObservableCollection<LinkWorksetsItem>();

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

            //сценарий работы плагина (открыть или закрыть выбранные РН)
            bool WSOpenClose;

            var inputWindow = new KR_WSofLinks(uidoc);
            bool? result = inputWindow.ShowDialog();

            //проверяем ввел ли пользователь имя РН, если нет, то отменяем работу плагина
            if (result == true)
            {
                LinksWorksetsSP = inputWindow.LinksWorksetsList;
                WSOpenClose = inputWindow.WorksetOpenClose;
            }
            else
            {
            // Пользователь отменил ввод
                return Result.Cancelled;
            }

            foreach (var linkInstance in linkInstances)
            {
                Document linkedDoc = linkInstance.GetLinkDocument();
                // если связь до запуска плагина выгружена или ссылка на связь не корректная, то пропускаем ее
                if (linkedDoc == null)
                    continue;

                // Получаем тип связи
                RevitLinkType linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                if (linkType == null) 
                    continue;

                //Получаем имя связи
                string linkName = linkType.Name;

                //Получаем путь у связи               
                ModelPath linkPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkedDoc.PathName);


                // Cписок имен открытых РН внутри связи
                var openedWorksetNames = new FilteredWorksetCollector(linkedDoc)
                                                                                .OfKind(WorksetKind.UserWorkset)
                                                                                .ToWorksets()
                                                                                .Where(ws => ws.IsOpen)
                                                                                .Select(ws => ws.Name)
                                                                                .ToHashSet();

                //Список всех рабочих наборов связи не открывая связь
                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(linkPath);
                //Список РН которые по итогу надо открыть
                IList<WorksetId> worksetsToOpen = new List<WorksetId>();

                foreach (var ws in worksets)
                {
                    //если выбран сценарий закрыть и соответственно галка не стоит
                    if (!WSOpenClose)
                    {
                        // Открываем только те, что не равны целевому и были открыты
                        if (!Cont(LinksWorksetsSP, linkName, ws.Name) && openedWorksetNames.Contains(ws.Name))
                        {
                            worksetsToOpen.Add(ws.Id);
                        }
                    }
                    //если выбран сценарий открыть и соответственно галка стоит
                    else
                    {
                        // Открываем только те, что были открыты и равны целевому                       
                        if (Cont(LinksWorksetsSP, linkName, ws.Name) || openedWorksetNames.Contains(ws.Name))
                        {
                            worksetsToOpen.Add(ws.Id);
                        }
                    }
                }

                // Конфигурация: по умолчанию все закрыты, открываем только нужные
                WorksetConfiguration wsConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                if (worksetsToOpen.Count > 0)
                    wsConfig.Open(worksetsToOpen);

                // Выгружаем связь
                linkType.Unload(null);

                // Перезагружаем связь
                linkType.LoadFrom(linkPath, wsConfig);
            }
            return Result.Succeeded;
        }

        /// <summary>
        /// Используется для проверки чекбокса опредленного РН в определенной связи
        /// </summary>
        public bool Cont(ObservableCollection<LinkWorksetsItem> sp, string lname, string targetWS)
        {
            foreach (LinkWorksetsItem link in sp)
            {
                if (link.LinkName == lname)
                {
                    foreach (WorksetItem workset in link.Worksets)
                    {
                        if (workset.Name == targetWS && workset.IsSelected == true)
                            return true;
                    }
                }
            }
            return false;
        }


    }
}
