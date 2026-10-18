namespace KPLN_Tools_OVVK.Forms.Entities
{
    public enum OVVK_SystemManagerRowKind
    {
        System,
        TypeParameterWarning,
        InvalidParameterError,
        EmptyParameterError,
        SystemNameError,
        ForbiddenSystemTypeError,
        SystemAbbreviationError,
        DuplicateSystemNameError,
    }

    public sealed class OVVK_SystemManagerRowM : System.ComponentModel.INotifyPropertyChanged
    {
        private bool _isSelected;
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                    return;
                _isSelected = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public OVVK_SystemManagerRowM(string systemName)
        {
            SystemName = systemName;
            DisplayName = systemName;
            Description = "Выделить элементы системы «" + systemName + "».";
        }

        internal OVVK_SystemManagerRowM(string systemName, string[] duplicateNames)
        {
            Kind = OVVK_SystemManagerRowKind.DuplicateSystemNameError;
            SystemName = systemName;
            string commonName = duplicateNames[0].Trim();
            DisplayName = systemName + " — ошибка: дубликат системы «" + commonName + "»";
            Description = "Дубликаты системы «" + commonName + "»: " + string.Join("; ", duplicateNames)
                + ".\nПриведи значения выбранного параметра к одному виду на всех элементах группы."
                + " До исправления всех дубликатов создание 3D-видов недоступно."
                + "\nНажми «Выделить элементы», чтобы найти элементы с этим значением.";
        }

        public OVVK_SystemManagerRowM(int typeParameterElementCount)
            : this(OVVK_SystemManagerRowKind.TypeParameterWarning, typeParameterElementCount)
        {
        }

        public OVVK_SystemManagerRowM(OVVK_SystemManagerRowKind kind, int elementCount, string details = null)
        {
            Kind = kind;
            string title;
            switch (kind)
            {
                case OVVK_SystemManagerRowKind.TypeParameterWarning:
                    title = "Предупреждение: параметр типа";
                    Description = "Значение общее для всех экземпляров типа. Для разных систем нужен параметр экземпляра.";
                    break;
                case OVVK_SystemManagerRowKind.InvalidParameterError:
                    title = "Ошибка: параметр отсутствует или не текстовый";
                    Description = "Выбранный параметр отсутствует у экземпляра и типа либо имеет неподходящий тип данных.";
                    break;
                case OVVK_SystemManagerRowKind.EmptyParameterError:
                    title = "Ошибка: пустое значение параметра";
                    Description = "Значение не задано (HasValue = false), пустое или состоит из пробелов. У вложенных семейств проверяется собственное значение.";
                    break;
                default:
                    title = "Ошибка: некорректное имя системы";
                    Description = "Вместо имени системы записано «Нет системы» или инструкция о ручном заполнении, Dynamo либо переносе значения.";
                    break;
            }

            DisplayName = $"{title} — элементов: {elementCount}";
            if (!string.IsNullOrWhiteSpace(details))
                Description += "\n" + details;
        }


        public string SystemName { get; }
        internal Autodesk.Revit.DB.ElementId SystemTypeId { get; }
        internal Autodesk.Revit.DB.ElementId[] ElementIds { get; }

        internal OVVK_SystemManagerRowM(OVVK_SystemManagerRowKind kind, Autodesk.Revit.DB.ElementId typeId,
            string typeName, string details, Autodesk.Revit.DB.ElementId[] elementIds)
        {
            Kind = kind;
            SystemTypeId = typeId;
            ElementIds = elementIds;
            string title = kind == OVVK_SystemManagerRowKind.ForbiddenSystemTypeError
                ? "нельзя использовать системные системы" : "неверное сокращение системы";
            DisplayName = $"Ошибка: {title} «{typeName}» — элементов: {elementIds.Length}";
            Description = details + "\nНажми «Выделить элементы», чтобы найти элементы этого типа системы.";
        }
        public string DisplayName { get; }
        public string Description { get; }
        public OVVK_SystemManagerRowKind Kind { get; }
        public bool IsTypeParameterWarning => Kind == OVVK_SystemManagerRowKind.TypeParameterWarning;
        public bool IsWarning => IsTypeParameterWarning;
        public bool IsError => Kind == OVVK_SystemManagerRowKind.InvalidParameterError
            || Kind == OVVK_SystemManagerRowKind.EmptyParameterError || Kind == OVVK_SystemManagerRowKind.SystemNameError
            || Kind == OVVK_SystemManagerRowKind.ForbiddenSystemTypeError || Kind == OVVK_SystemManagerRowKind.SystemAbbreviationError
            || Kind == OVVK_SystemManagerRowKind.DuplicateSystemNameError;
        public bool IsIssue => IsError || IsWarning;
    }
}
