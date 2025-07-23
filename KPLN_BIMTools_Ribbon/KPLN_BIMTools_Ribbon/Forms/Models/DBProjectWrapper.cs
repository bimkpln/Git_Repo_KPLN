using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_BIMTools_Ribbon.Forms.Models
{
    public class DBProjectWrapper : INotifyPropertyChanged, IDataErrorInfo
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _wrName;
        private string _wrCode;
        private string _wrStage;
        private int _wrRevitVersion = 2023;
        public string _wrServerPath;
        public string _wrRevitServerPath;
        public string _wrRevitServerPath2;
        public string _wrRevitServerPath3;
        public string _wrRevitServerPath4;
        public bool _wrIsClosed = false;

        public DBProjectWrapper()
        {
            // Лучше такое из БД брать, но оно 99,9% статично
            WrStageColl = new string[]
            {
                "АФК",
                "ПД",
                "ПД_Корр",
                "РД",
                "АН",
            };
            
            WrStage = WrStageColl[0];
        }

        /// <summary>
        /// Имя проекта
        /// </summary>
        public string WrName
        {
            get => _wrName;
            set
            {
                _wrName = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Код (аббревиатура) проекта
        /// </summary>
        public string WrCode
        {
            get => _wrCode;
            set
            {
                _wrCode = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Стадия проектирования
        /// </summary>
        public string WrStage
        {
            get => _wrStage;
            set
            {
                _wrStage = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Коллекция возможных стадий проектирования
        /// </summary>
        public string[] WrStageColl { get; }

        /// <summary>
        /// Версия используемого Revit
        /// </summary>
        public int WrRevitVersion
        {
            get => _wrRevitVersion;
            set
            {
                _wrRevitVersion = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Путь к корневой папке
        /// </summary>
        public string WrServerPath
        {
            get => _wrServerPath;
            set
            {
                _wrServerPath = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Путь Revit-Server
        /// </summary>
        public string WrRevitServerPath
        {
            get => _wrRevitServerPath;
            set
            {
                _wrRevitServerPath = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Путь Revit-Server №2
        /// </summary>
        public string WrRevitServerPath2
        {
            get => _wrRevitServerPath2;
            set
            {
                _wrRevitServerPath2 = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Путь Revit-Server №3
        /// </summary>
        public string WrRevitServerPath3
        {
            get => _wrRevitServerPath3;
            set
            {
                _wrRevitServerPath3 = value;
                OnPropertyChanged();
            }
        }
        /// <summary>
        /// Путь Revit-Server №4
        /// </summary>
        public string WrRevitServerPath4
        {
            get => _wrRevitServerPath4;
            set
            {
                _wrRevitServerPath4 = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Режим блокировки проекта под набор разрешенных пользователей вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool WrIsClosed
        {
            get => _wrIsClosed;
            set
            {
                _wrIsClosed = value;
                OnPropertyChanged();
            }
        }

        public string Error => null;

        public string this[string columnName]
        {
            get
            {
                switch (columnName)
                {
                    case nameof(WrServerPath):
                        return ValidateServerPath(WrServerPath, @"Y:\", '/', '\\', "сервер КПЛН");
                    case nameof(WrRevitServerPath):
                        return ValidateServerPath(WrRevitServerPath, @"RSN://", '\\', '/', "Revit-Server КПЛН");
                    case nameof(WrRevitServerPath2):
                        return ValidateServerPath(WrRevitServerPath2, @"RSN://", '\\', '/', "Revit-Server КПЛН");
                    case nameof(WrRevitServerPath3):
                        return ValidateServerPath(WrRevitServerPath3, @"RSN://", '\\', '/', "Revit-Server КПЛН");
                    case nameof(WrRevitServerPath4):
                        return ValidateServerPath(WrRevitServerPath4, @"RSN://", '\\', '/', "Revit-Server КПЛН");
                }

                return null;
            }
        }

        private string ValidateServerPath(string value, string start, char wrongSlash, char rightSlash, string name)
        {
            if (string.IsNullOrEmpty(value))
                return null;
            if (!value.StartsWith(start))
                return $"Путь на {name} должен начинаться с '{start}'";
            if (value.Contains(wrongSlash))
                return $"Путь на {name} должен разделяться символом '{rightSlash}'";
            return null;
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
