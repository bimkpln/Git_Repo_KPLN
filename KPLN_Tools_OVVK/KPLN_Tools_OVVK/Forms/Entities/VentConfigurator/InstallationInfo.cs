using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class InstallationInfo : EditableValue
    {
        internal const string SystemNameParameter = "КП_О_Имя Системы";
        internal const string SystemNameTypeParameter = "Имя системы_Тип";
        internal const string SystemNameByTypeParameter = "Имя системы_По типу";
        internal const string SystemNameInstanceParameter = "Имя системы_Экземпляр";
        internal const string ManufacturerParameter = "КП_О_Завод-изготовитель";
        internal const string MarkParameter = "КП_О_Марка";
        internal const string UnitParameter = "КП_О_Единица измерения";
        internal const string DescriptionParameter = "КП_О_Наименование";
        internal const string ProductCodeParameter = "КП_О_Код изделия", MassTextParameter = "КП_О_Масса_Текст";
        private string _systemName = "", _manufacturer = "", _mark = "", _unit = "", _description = "", _productCode = "", _massText = "";

        private void Set(ref string field, string value, [CallerMemberName] string property = null)
        {
            value = value ?? "";
            if (field == value)
                return;
            field = value;
            Changed(property);
        }

        public string SystemName
        {
            get
            {
                return _systemName;
            }
            set
            {
                Set(ref _systemName, value);
            }
        }

        public string Manufacturer
        {
            get
            {
                return _manufacturer;
            }
            set
            {
                Set(ref _manufacturer, value);
            }
        }

        public string Mark
        {
            get
            {
                return _mark;
            }
            set
            {
                Set(ref _mark, value);
            }
        }

        public string Unit
        {
            get
            {
                return _unit;
            }
            set
            {
                Set(ref _unit, value);
            }
        }

        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                Set(ref _description, value);
            }
        }

        public string ProductCode
        {
            get
            {
                return _productCode;
            }
            set
            {
                Set(ref _productCode, value);
            }
        }

        public string MassText
        {
            get
            {
                return _massText;
            }
            set
            {
                Set(ref _massText, value);
            }
        }

        internal string AutomaticTypeName
        {
            get
            {
                return string.IsNullOrWhiteSpace(SystemName) || string.IsNullOrWhiteSpace(Mark)
            ? "Новый тип" : SystemName.Trim() + "_" + Mark.Trim();
            }
        }

        internal Dictionary<string, string> Assignments()
        {
            return new Dictionary<string, string> { { SystemNameParameter, SystemName }, { ManufacturerParameter, Manufacturer },
                { MarkParameter, Mark }, { UnitParameter, Unit }, { DescriptionParameter, Description },
                { ProductCodeParameter, ProductCode }, { MassTextParameter, MassText } };
        }

        internal void SetParameterValue(string name, string value)
        {
            switch (name)
            {
                case SystemNameParameter:
                    SystemName = value;
                    break;
                case ManufacturerParameter:
                    Manufacturer = value;
                    break;
                case MarkParameter:
                    Mark = value;
                    break;
                case UnitParameter:
                    Unit = value;
                    break;
                case DescriptionParameter:
                    Description = value;
                    break;
                case ProductCodeParameter:
                    ProductCode = value;
                    break;
                case MassTextParameter:
                    MassText = value;
                    break;
            }
        }

        internal InstallationInfo Copy()
        {
            return new InstallationInfo
            {
                SystemName = SystemName,
                Manufacturer = Manufacturer,
                Mark = Mark,
                Unit = Unit,
                Description = Description,
                ProductCode = ProductCode,
                MassText = MassText
            };
        }
    }
}
