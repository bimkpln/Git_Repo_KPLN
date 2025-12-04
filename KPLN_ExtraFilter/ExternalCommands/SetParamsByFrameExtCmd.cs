using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    /// <summary>
    /// Класс фильтрации Selection 
    /// </summary>
    internal class SelectorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Фильтрация по классу (проемы)
            if (elem is Opening _)
                return false;

            // Фильтрация по категории (модельные эл-ты, кроме видов, сборок)
            if (elem.Category is Category elCat)
            {
                int elCatId = elCat.Id.IntegerValue;
                if (((elem.Category.CategoryType == CategoryType.Model)
                        || (elem.Category.CategoryType == CategoryType.Internal))
                    && (elCatId != (int)BuiltInCategory.OST_Viewers)
                    && (elCatId != (int)BuiltInCategory.OST_IOSModelGroups)
                    && (elCatId != (int)BuiltInCategory.OST_Assemblies))
                    return true;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SetParamsByFrameExtCmd : IExternalCommand
    {
        internal const string PluginName = "Выбрать/заполнить рамкой";
        private SetParamsByFrameForm _mainForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Пользовательский элемент
            IEnumerable<Element> userSelElems = null;
            var userSelId = uidoc.Selection.GetElementIds();
            if (userSelId.Any())
                userSelElems = userSelId.Select(id => doc.GetElement(id));

            _mainForm = new SetParamsByFrameForm(doc, userSelElems);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(_mainForm);
            // Предустановка
            if (userSelElems != null)
            {
                _mainForm.CurrentSetParamsByFrameVM.CurrentSetParamsByFrameM.Doc = doc;
                _mainForm.CurrentSetParamsByFrameVM.CurrentSetParamsByFrameM.UserSelElems = userSelElems;
            }


            // Создание ExternalEvent для переключения видов
            ViewActivatedHandler viewHandler = new ViewActivatedHandler();
            ExternalEvent viewExtEv = ExternalEvent.Create(viewHandler);

            // Создание ExternalEvent для отписки от переключения видов (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubViewActHandler unsubViewHandler = new UnsubViewActHandler() { Handler = OnViewChanged };
            ExternalEvent unsubViewExtEv = ExternalEvent.Create(unsubViewHandler);

#if !Debug2020 && !Revit2020
            // Создание ExternalEvent для выделения эл-в
            SelectionChangedHandler selHandler = new SelectionChangedHandler();
            ExternalEvent selExtEv = ExternalEvent.Create(selHandler);

            // Создание ExternalEvent для отписки от выбора эл-в (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubSelChHandler unsubSelHandler = new UnsubSelChHandler() { Handler = OnSelectionChanged };
            ExternalEvent unsubSelExtEv = ExternalEvent.Create(unsubSelHandler);

            // Доп настройки окна
            _mainForm.SetExternalEvent(viewExtEv, viewHandler, selExtEv, selHandler);
#endif

            // Подписываюсь на OnViewChanged
            uiapp.ViewActivated += OnViewChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubViewExtEv.Raise();

#if !Debug2020 && !Revit2020
            // Подписываюсь на SelectionChanged
            uiapp.SelectionChanged += OnSelectionChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubSelExtEv.Raise();
#endif


            _mainForm.Show();


            return Result.Succeeded;
        }

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => _mainForm.RaiseUpdateSelChanged();
#endif

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => _mainForm.RaiseUpdateViewChanged();
    }
}
