using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Commands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_ConfigWorker;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms.ViewModels
{
    public sealed class SelectionByModelVM
    {
        private readonly SelectionByModel _mainWindow;

        public SelectionByModelVM(SelectionByModel mainWindow, Document doc)
        {
            _mainWindow = mainWindow;
            // Чтение конфигурации последнего запуска
            object lastRunConfigObj = ConfigService.ReadConfigFile<SelectionByClickM>(ModuleData.RevitVersion, doc, ConfigType.Memory);
            if (lastRunConfigObj != null && lastRunConfigObj is SelectionByModelM model)
            {
                // Если тот же док - беру из пред. запуска
                if (model.Doc.Title == doc.Title)
                    CurrentSelectionByModelM = model;
            }

            if (CurrentSelectionByModelM == null)
                CurrentSelectionByModelM = new SelectionByModelM(doc);


            DropUserParamCmd = new RelayCommand<object>(_ => DropUserParam());
            ModelSelectionCmd = new RelayCommand<object>(_ => ModelSelection());
            UpdateTreeElemsCmd = new RelayCommand<object>(_ => UpdateTreeElems());

            AddCategoryCmd = new RelayCommand<object>(_ => AddCategory());
            RemoveCategoryCmd = new RelayCommand<SelectionByModelM_CategoryM>(RemoveCategory);
            AddParameterCmd = new RelayCommand<object>(_ => AddParameter());
            RemoveParameterCmd = new RelayCommand<SelectionByModelM_ParamM>(RemoveParameter);

            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public SelectionByModelM CurrentSelectionByModelM { get; set; }

        /// <summary>
        /// Комманда: Создать новую выборку в модели
        /// </summary>
        public ICommand ModelSelectionCmd { get; }

        /// <summary>
        /// Комманда: Сброс выбора
        /// </summary>
        public ICommand DropUserParamCmd { get; }

        /// <summary>
        /// Комманда: Обновить элементы в дереве
        /// </summary>
        public ICommand UpdateTreeElemsCmd { get; }

        /// <summary>
        /// Комманда: Добавить параметр фильтрации
        /// </summary>
        public ICommand AddCategoryCmd { get; }

        /// <summary>
        /// Комманда: Удалить параметр фильтрации
        /// </summary>
        public ICommand RemoveCategoryCmd { get; }

        /// <summary>
        /// Комманда: Добавить категорию группирования
        /// </summary>
        public ICommand AddParameterCmd { get; }

        /// <summary>
        /// Комманда: Удалить категорию группирования
        /// </summary>
        public ICommand RemoveParameterCmd { get; }

        /// <summary>
        /// Комманда: Добавить выборку к уже существующей выборке в модели
        /// </summary>
        public ICommand AddModelSelectionCmd { get; }

        /// <summary>
        /// Комманда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void DropUserParam() => CurrentSelectionByModelM.DropToDefault();

        public void ModelSelection() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectionByModelExcCmd(CurrentSelectionByModelM));

        public void UpdateTreeElems() => CurrentSelectionByModelM.SetUserSelElems();

        public void AddCategory()
        {
            if (CurrentSelectionByModelM.Where_SelectedCategories.Count < 3)
            {
                if (CurrentSelectionByModelM.Where_SelectedCategories.Any(c => c.CatM_SelectedCategory == null))
                    MessageBox.Show(_mainWindow, "Добавить новую категорию можно только если все ранее добавлены задействованы (сейчасесть пустые поля)", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                {
                    CurrentSelectionByModelM.Where_SelectedCategories.Add(new SelectionByModelM_CategoryM(CurrentSelectionByModelM));
                    CurrentSelectionByModelM.UpdateCanRunANDUserHelp();
                }
            }
            else
                MessageBox.Show(_mainWindow, "Максимально фильтровать можно по 3 категориям.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        public void RemoveCategory(SelectionByModelM_CategoryM item)
        {
            if (CurrentSelectionByModelM.Where_SelectedCategories.Count == 1)
                MessageBox.Show(_mainWindow, "Нельзя удалить последнюю категорию из списка. Если НЕ нужно фильтровать по категориям - сними галку с параметра \"Только этой/-их категории/-ий:\".", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            else if (item is SelectionByModelM_CategoryM c)
            {
                string catName = c.CatM_SelectedCategory == null ? "Пустая категория" : c.CatM_SelectedCategory.RevitCatName;

                var td = MessageBox.Show(_mainWindow, $"Сейчас из фильтров будет удалена категория \"{catName}\"", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (td == MessageBoxResult.Yes)
                {
                    CurrentSelectionByModelM.Where_SelectedCategories.Remove(c);
                    
                    CurrentSelectionByModelM.SetUserSelElems();
                    CurrentSelectionByModelM.UpdateCanRunANDUserHelp();
                }
            }
        }

        public void AddParameter()
        {
            if (CurrentSelectionByModelM.What_SelectedParameters.Count < 3)
            {
                if (CurrentSelectionByModelM.What_SelectedParameters.Any(c => c.ParamM_SelectedParameter == null))
                    MessageBox.Show(_mainWindow, "Добавить новый параметр можно только если все ранее добавлены задействованы (сейчасесть пустые поля)", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                {
                    CurrentSelectionByModelM.What_SelectedParameters.Add(new SelectionByModelM_ParamM(CurrentSelectionByModelM));
                    CurrentSelectionByModelM.UpdateCanRunANDUserHelp();
                }
            }
            else
                MessageBox.Show(_mainWindow, "Максимально группировать можно по 3 параметрам.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        public void RemoveParameter(SelectionByModelM_ParamM item)
        {
            if (CurrentSelectionByModelM.What_SelectedParameters.Count == 1)
                MessageBox.Show(_mainWindow, "Нельзя удалить последний параметр из списка. Если НЕ нужно группировать по параметрам - сними галку с параметра \"Одинакового значения пар-ра\".", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            else if (item is SelectionByModelM_ParamM c)
            {
                string paramName = c.ParamM_SelectedParameter == null ? "Пустой параметр" : c.ParamM_SelectedParameter.RevitParamName;

                var td = MessageBox.Show(_mainWindow, $"Сейчас из фильтров будет удален параметр \"{paramName}\"", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (td == MessageBoxResult.Yes)
                {
                    CurrentSelectionByModelM.What_SelectedParameters.Remove(c);
                    
                    CurrentSelectionByModelM.SetUserSelElems();
                    CurrentSelectionByModelM.UpdateCanRunANDUserHelp();
                }
            }
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
