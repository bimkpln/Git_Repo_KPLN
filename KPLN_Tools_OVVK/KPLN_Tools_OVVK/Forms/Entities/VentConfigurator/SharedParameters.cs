using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal static class SharedParameters
    {
        internal const string FrameVisible = "Рама_Видимость", ServiceRight = "ЗО_Справа", ServiceDepth = "ЗО_Глубина";
        internal const string HeaterRight = "Нагреватель_Трубки_Справа", CoolerRight = "Охладитель_Трубки_Справа";
        internal static readonly string[] WaterLengths = {
                "Вода_Патрубок_Обратка_Смещение X", "Вода_Патрубок_Обратка_Смещение Y",
                "Вода_Патрубок_Приток/Обратка_Диаметр", "Вода_Патрубок_Приток_Смещение X", "Вода_Патрубок_Приток_Смещение Y" };
        internal static readonly string[] DrainLengths = {
                "Дренаж_Патрубок_Диаметр", "Дренаж_Патрубок_Смещение X", "Дренаж_Патрубок_Смещение Y" };
        internal static readonly string[] RefrigerantLengths = {
                "Фреон_Патрубок_Обратка_Диаметр", "Фреон_Патрубок_Обратка_Смещение X", "Фреон_Патрубок_Обратка_Смещение Y",
                "Фреон_Патрубок_Приток_Диаметр", "Фреон_Патрубок_Приток_Смещение X", "Фреон_Патрубок_Приток_Смещение Y" };

        internal static IEnumerable<string> LengthNames
        {
            get
            {
                return new[] { ServiceDepth }.Concat(WaterLengths).Concat(DrainLengths).Concat(RefrigerantLengths);
            }
        }

        internal static readonly string[] BooleanNames = { FrameVisible, ServiceRight, HeaterRight, CoolerRight };

        internal static IEnumerable<string> ForSection(SectionTypeChoice type)
        {
            var kind = type == null ? ConnectionKind.None : type.Connections;
            if (kind == ConnectionKind.WaterHeater)
                return new[] { HeaterRight }.Concat(WaterLengths);
            if (kind == ConnectionKind.WaterCooler)
                return new[] { CoolerRight }.Concat(WaterLengths).Concat(DrainLengths);
            if (kind == ConnectionKind.RefrigerantCooler)
                return new[] { CoolerRight }.Concat(DrainLengths).Concat(RefrigerantLengths);
            return Enumerable.Empty<string>();
        }

        internal static string Label(string name)
        {
            if (name == FrameVisible)
                return "Рама";
            if (name == ServiceRight)
                return "Зона обслуживания";
            if (name == ServiceDepth)
                return "Зона обслуживания: глубина";
            if (name == HeaterRight || name == CoolerRight)
                return "Сторона трубок";
            return name.Replace("_Патрубок_", ": ").Replace('_', ' ');
        }
    }
}
