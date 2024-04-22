using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Runtime.CompilerServices;

namespace KPLN_BIMTools_Ribbon.Core.SQLite.Entities
{
    public class DBConfigEntity : INotifyPropertyChanged
    {
        private string _name;
        private string _pathFrom;
        private string _pathTo;

        /// <summary>
        /// Для работы Dapper нужен пустой конструктор
        /// </summary>
        public DBConfigEntity()
        {
        }

        public DBConfigEntity(string name, string pathFrom, string pathTo) : this()
        {
            Name = name;
            PathFrom = pathFrom;
            PathTo = pathTo;
        }

        [Key]
        public int Id { get; set; }

        public string Name
        {
            get => _name;
            set { SetField(ref _name, value); }
        }

        /// <summary>
        /// Путь к файлу ДЛЯ обработки
        /// </summary>
        public string PathFrom
        {
            get => _pathFrom;
            set { SetField(ref _pathFrom, value); }
        }

        /// <summary>
        /// Путь к итоговому файлу ПОСЛЕ обработки
        /// </summary>
        public string PathTo
        {
            get => _pathTo;
            set { SetField(ref _pathTo, value); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // Универсальный метод для установки значения поля и вызова события PropertyChanged
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);

            return true;
        }
    }
}
