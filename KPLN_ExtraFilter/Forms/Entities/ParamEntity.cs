using Autodesk.Revit.DB;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для параметра
    /// </summary>
    public sealed class ParamEntity
    {
        public ParamEntity()
        {
        }

        public ParamEntity(Parameter param)
        {
            RevitParamName = param.Definition.Name;
            RevitParamIntId = param.Id.IntegerValue;
        }

        public ParamEntity(Parameter param, string tooltip) : this (param)
        {
            CurrentToolTip = tooltip;
        }

        /// <summary>
        /// Имя параметра
        /// </summary>
        public string RevitParamName { get; set; }

        /// <summary>
        /// ID параметра
        /// </summary>
        public int RevitParamIntId { get; set; }

        /// <summary>
        /// Дополнительное описание пар-ра для wpf
        /// </summary>
        public string CurrentToolTip { get; set; }

        /// <summary>
        /// Переопределение метода Equals. ОБЯЗАТЕЛЬНО для десериализации, т.к. иначе на wpf не может найти эквивалетный инстанс
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj is ParamEntity other)
                return RevitParamName == other.RevitParamName && RevitParamIntId == other.RevitParamIntId;

            return false;
        }

        /// <summary>
        /// Переопределение метода GetHashCode
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            // Используем простое XOR-сочетание хэш-кодов свойств
            return RevitParamName.GetHashCode() ^ RevitParamIntId.GetHashCode();
        }
    }
}
