using Autodesk.Revit.DB;

namespace KPLN_Publication.Common
{
    public class ListBoxParameter
    {
        public string Name { get; set; }
        public string Group { get; set; }
        public string ToolTip { get; set; }
        public Parameter Parameter { get; set; }
        public ListBoxParameter(Parameter p, string tooltip)
        {
            //Print("ListBoxParameter", KPLN_Loader.Preferences.MessageType.System_OK);
            Parameter = p;
            Name = p.Definition.Name;
            Group = LabelUtils.GetLabelFor(p.Definition.ParameterGroup);
            if (p.IsShared)
            { ToolTip = string.Format("<{0}> - Общий параметр", tooltip); }
            else
            { ToolTip = string.Format("<{0}> - Встроенный параметр", tooltip); }
        }
        public ListBoxParameter()
        {
            //Print("ListBoxParameter", KPLN_Loader.Preferences.MessageType.System_OK);
            Parameter = null;
            Name = "<Пусто>";
            Group = "<Нет>";
            ToolTip = "Не выбрано";
        }
    }
}
