using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
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

            // Анбоксинг на View Revit
            if (elem is View view)
                Name = view.Title;

            // Анбоксинг на DbProject KPLN_Library_DataBase
            else if (elem is DBProject dbProject)
            {
                if (dbProject.IsClosed)
                    Name = $"(АРХИВ){dbProject.Name}(АРХИВ). Стадия: {dbProject.Stage}";
                else
                    Name = $"{dbProject.Name}. Стадия: {dbProject.Stage}";
            }

            // Анбоксинг на Element Revit
            else if (elem is Element element)
                Name = element.Name;

            // Анбоксинг на Parameter Revit
            else if (elem is Parameter param)
                Name = param.Definition.Name;

            // Анбоксинг на string (Element == Name, например - путь к файлу)
            else if (elem is string name)
                Name = name;
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

        public override bool Equals(object obj)
        {
            if (obj is ElementEntity otherEntity)
            {
                return string.Equals(this.Name, otherEntity.Name, StringComparison.Ordinal);
            }
            return false;
        }

        public override int GetHashCode() => this.Name?.GetHashCode() ?? 0;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
