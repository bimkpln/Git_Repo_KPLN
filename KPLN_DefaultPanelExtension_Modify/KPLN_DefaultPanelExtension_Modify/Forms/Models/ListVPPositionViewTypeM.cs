using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class ListVPPositionViewTypeM
    {
        public const string UnknownCode = "Неизв";

        private ListVPPositionViewTypeM(string code, string name)
        {
            Code = code;
            Name = name;
        }

        public string Code { get; }

        public string Name { get; }

        public static IReadOnlyList<ListVPPositionViewTypeM> Items { get; } = new[]
        {
            new ListVPPositionViewTypeM(string.Empty, "Все типы"),
            new ListVPPositionViewTypeM("ПП", "ПП — План этажа"),
            new ListVPPositionViewTypeM("Р", "Р — Разрез"),
            new ListVPPositionViewTypeM("3D", "3D — 3D-вид"),
            new ListVPPositionViewTypeM("СП", "СП — Спецификация"),
            new ListVPPositionViewTypeM("Л", "Л — Легенда"),
            new ListVPPositionViewTypeM("ЧВ", "ЧВ — Чертёжный вид"),
            new ListVPPositionViewTypeM(UnknownCode, "Неизв — Остальные / без типа"),
        };

        public static string GetCode(Element element)
        {
            if (element is ScheduleSheetInstance)
                return "СП";

            View view = element as View;
            if (element is Viewport viewport)
                view = viewport.Document.GetElement(viewport.ViewId) as View;

            return view == null ? UnknownCode : GetCode(view.ViewType);
        }

        public static string GetCode(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.FloorPlan: return "ПП";
                case ViewType.Section: return "Р";
                case ViewType.ThreeD: return "3D";
                case ViewType.Schedule: return "СП";
                case ViewType.Legend: return "Л";
                case ViewType.DraftingView: return "ЧВ";
                default: return UnknownCode;
            }
        }

        public static string NormalizeCode(string code) =>
            Items.Any(item => item.Code.Length > 0 && item.Code == code) ? code : UnknownCode;
    }
}
