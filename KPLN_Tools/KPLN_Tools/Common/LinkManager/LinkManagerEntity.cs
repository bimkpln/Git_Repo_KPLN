using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.LinkManager
{
    [Serializable]
    public class LinkManagerEntity : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly CoordinateType[] _linkCoordinateTypeColl = new CoordinateType[]
        {
            new CoordinateType("Авто - По общим координатам", ImportPlacement.Shared),
            new CoordinateType("Авто - Совмещение внутренних начал", ImportPlacement.Origin),
        };
        private string _linkName;
        private string _linkPath;
        private CoordinateType _linkCoordinateType;
        private string _worksetToCloseNamesStartWith;
        private bool _createWorksetForLinkInst = true;

        [JsonConstructor]
        public LinkManagerEntity()
        {
        }

        public LinkManagerEntity(string name, string path)
        {
            LinkName = name;
            LinkPath = path;
            LinkCoordinateType = LinkCoordinateTypeColl[0];
        }

        /// <summary>
        /// Имя связи
        /// </summary>
        public string LinkName
        {
            get => _linkName;
            set
            {
                _linkName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Путь к связи
        /// </summary>
        public string LinkPath
        {
            get => _linkPath;
            set
            {
                _linkPath = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Типы плошадок
        /// </summary>
        public CoordinateType[] LinkCoordinateTypeColl
        {
            get => _linkCoordinateTypeColl;
        }

        /// <summary>
        /// Выбранный тип плошадки
        /// </summary>
        public CoordinateType LinkCoordinateType
        {
            get => _linkCoordinateType;
            set
            {
                _linkCoordinateType = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Создать отдельный рабочий набор для связи?
        /// </summary>
        public bool CreateWorksetForLinkInst
        {
            get => _createWorksetForLinkInst;
            set
            {
                _createWorksetForLinkInst = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Имя рабочих наборов, которые нужно закрыть (имя начинается с)
        /// </summary>
        public string WorksetToCloseNamesStartWith
        {
            get => _worksetToCloseNamesStartWith;
            set
            {
                _worksetToCloseNamesStartWith = value;
                NotifyPropertyChanged();
            }
        }

        public object ToJson()
        {
            return new
            {
                this.LinkName,
                this.LinkPath,
                this.LinkCoordinateType,
                this.CreateWorksetForLinkInst,
                this.WorksetToCloseNamesStartWith,
            };
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
