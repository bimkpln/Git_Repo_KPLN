using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_Tools.Common.LinkManager
{
    [Serializable]
    public class LinkManagerUpdateEntity : LinkManagerEntity, IJsonSerializable
    {
        private string _updatedLinkName;
        private string _updatedLinkPath;

        [JsonConstructor]
        public LinkManagerUpdateEntity() : base() { }

        public LinkManagerUpdateEntity(string name, string path, string updatedName, string updatedPath, EntityStatus status = EntityStatus.Ok) : base(name, path) 
        {
            UpdatedLinkName = updatedName;
            UpdatedLinkPath = updatedPath;

            CurrentEntStatus = status;
        }

        /// <summary>
        /// Обновленное имя связи
        /// </summary>
        public string UpdatedLinkName
        {
            get => _updatedLinkName;
            set
            {
                _updatedLinkName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Обновленный путь к связи
        /// </summary>
        public string UpdatedLinkPath
        {
            get => _updatedLinkPath;
            set
            {
                _updatedLinkPath = value;
                NotifyPropertyChanged();
            }
        }

        public object ToJson()
        {
            return new
            {
                this.LinkName,
                this.LinkPath,
                this.UpdatedLinkName,
                this.UpdatedLinkPath,
            };
        }

    }

}
