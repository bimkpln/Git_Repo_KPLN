using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Library_Forms.Common
{
    public class ElementEntity : INotifyPropertyChanged
    {
        private string _name;
        private string _tooltip;
        private bool _isSelected = false;

        public event PropertyChangedEventHandler PropertyChanged;

        public ElementEntity(object elem)
        {
            Element = elem;

            // Анбоксинг на Element Revit
            if (elem is Element element)
                Name = element.Name;

            // Анбоксинг на DbProject KPLN_Library_DataBase
            if (elem is DBProject dbProject)
                Name = $"{dbProject.Name}. Стадия: {dbProject.Stage}";
        }

        public ElementEntity(object elem, string tooltip) : this(elem)
        {
            Tooltip = tooltip;
        }

        /// <summary>
        /// Элемент для отображения в окне выбора
        /// </summary>
        public object Element { get; private set; }

        /// <summary>
        /// Имя элемента, отображаемое пользователю
        /// </summary>
        public string Name 
        {
            get => _name;
            private set
            {
                if (value != _name)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Подсказка для пользователя
        /// </summary>
        public string Tooltip
        {
            get => _tooltip;
            private set
            {
                if (value != _tooltip)
                {
                    _tooltip = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Для окна wpf - пометка, что элемент выбран
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
