using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using KPLN_Tools.Forms.Models.Core;
using Newtonsoft.Json;
using System;

namespace KPLN_Tools.Common.LinkManager
{
    [Serializable]
    public class LinkManagerLoadEntity : LinkManagerEntity, IJsonSerializable
    {
        private readonly CoordinateType[] _linkCoordinateTypeColl = new CoordinateType[]
        {
            new CoordinateType("Авто - По общим координатам", ImportPlacement.Shared),
            new CoordinateType("Авто - Совмещение внутренних начал", ImportPlacement.Origin),
        };

        private CoordinateType _linkCoordinateType;
        private string _worksetToCloseNamesStartWith;
        private bool _createWorksetForLinkInst = true;

        [JsonConstructor]
        public LinkManagerLoadEntity() : base() { }

        public LinkManagerLoadEntity(string name, string path) : base(name, path)
        {
            LinkCoordinateType = LinkCoordinateTypeColl[0];
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
    }
}
