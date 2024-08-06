using Autodesk.Revit.DB;

namespace KPLN_Tools.Common.LinkManager
{
    public class CoordinateType
    {
        public CoordinateType(string name, ImportPlacement type)
        {
            Name = name;
            Type = type;
        }

        public string Name { get; private set; }

        public ImportPlacement Type { get; private set; }

        // Переопределение метода Equals. ОБЯЗАТЕЛЬНО для десериализации, т.к. иначе на wpf не может найти эквивалетный инстанс
        public override bool Equals(object obj)
        {
            if (obj is CoordinateType other)
            {
                return Name == other.Name && Type == other.Type;
            }
            return false;
        }

        // Переопределение метода GetHashCode
        public override int GetHashCode()
        {
            // Используем простое XOR-сочетание хэш-кодов свойств
            return Name.GetHashCode() ^ Type.GetHashCode();
        }
    }
}
