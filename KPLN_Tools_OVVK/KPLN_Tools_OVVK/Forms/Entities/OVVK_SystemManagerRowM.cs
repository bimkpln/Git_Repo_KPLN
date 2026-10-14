namespace KPLN_Tools_OVVK.Forms.Entities
{
    public enum OVVK_SystemManagerRowKind
    {
        System,
        TypeParameterWarning,
        InvalidParameterError,
        EmptyParameterError,
        SystemNameError,
    }

    public sealed class OVVK_SystemManagerRowM
    {
        public OVVK_SystemManagerRowM(string systemName)
        {
            SystemName = systemName;
            DisplayName = systemName;
            Description = "Выделить элементы системы «" + systemName + "».";
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
        public string DisplayName { get; }
        public string Description { get; }
        public OVVK_SystemManagerRowKind Kind { get; }
        public bool IsTypeParameterWarning => Kind == OVVK_SystemManagerRowKind.TypeParameterWarning;
        public bool IsWarning => IsTypeParameterWarning;
        public bool IsError => Kind == OVVK_SystemManagerRowKind.InvalidParameterError
            || Kind == OVVK_SystemManagerRowKind.EmptyParameterError || Kind == OVVK_SystemManagerRowKind.SystemNameError;
        public bool IsIssue => IsError || IsWarning;
    }
}
