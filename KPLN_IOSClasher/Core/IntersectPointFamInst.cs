using Autodesk.Revit.DB;
using System;
using System.Globalization;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность для работы с точками пересечений
    /// </summary>
    internal class IntersectPointFamInst
    {
        public static readonly string ClashPointFamilyName = "ClashPoint_Small";

        public static readonly Guid PointCoord_Param = new Guid("76b464e2-b02e-4af8-bd1a-eb4985a0f362");
        public static readonly Guid AddedElementId_Param = new Guid("bf5ddf52-e83e-4094-9baf-db5686356b5d");
        public static readonly Guid OldElementId_Param = new Guid("dc3f91c3-8144-4e2c-9b4f-ea060706a2ea");
        public static readonly Guid LinkInstanceId_Param = new Guid("468b7905-8f4e-4f72-86f3-3a18685a3838");
        public static readonly Guid UserData_Param = new Guid("41b103e9-3eb8-43fe-af9d-b74a13acf45b");
        public static readonly Guid CurrentData_Param = new Guid("8984f2e6-606d-4cd0-8fb7-19bb034c5059");

        protected static XYZ ParseStringToXYZ(string input)
        {
            try
            {
                // Выдаляем дужкі і разбіваем радок па косцы
                string cleanedInput = input.Trim('(', ')');
                string[] parts = cleanedInput.Split(',');

                // Парсим кожную частку ў double
                double x = double.Parse(parts[0], CultureInfo.InvariantCulture);
                double y = double.Parse(parts[1], CultureInfo.InvariantCulture);
                double z = double.Parse(parts[2], CultureInfo.InvariantCulture);

                // Ствараем і вяртаем XYZ
                return new XYZ(x, y, z);
            }
            catch (Exception ex)
            {
                // Абработка памылак парсінгу
                throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из строки в XYZ", ex);
            }
        }
    }
}
