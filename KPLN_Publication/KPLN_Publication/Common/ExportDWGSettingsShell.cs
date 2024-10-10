using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace KPLN_Publication.Common
{
    /// <summary>
    /// Оболочка для возможности сериализации
    /// </summary>
    public class ExportDWGSettingsShell
    {
        public ExportDWGSettingsShell()
        {
        }

        public ExportDWGSettingsShell(string name, ExportDWGSettings setting)
        {
            Name = name;
            DWGExportSetting = setting;
        }

        public string Name { get; set; }

        [XmlIgnore]
        public ExportDWGSettings DWGExportSetting { get; private set; }

        // Переопределение метода Equals. ОБЯЗАТЕЛЬНО для десериализации, т.к. иначе на wpf не может найти эквивалетный инстанс
        public override bool Equals(object obj)
        {
            if (obj is ExportDWGSettings other)
            {
                return Name == other.Name;
            }
            return false;
        }

        // Переопределение метода GetHashCode
        public override int GetHashCode()
        {
            // Используем простое XOR-сочетание хэш-кодов свойств
            return Name.GetHashCode() ^ DWGExportSetting.GetHashCode();
        }
    }
}
