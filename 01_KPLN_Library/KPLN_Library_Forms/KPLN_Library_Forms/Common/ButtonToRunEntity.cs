namespace KPLN_Library_Forms.Common
{
    public class ButtonToRunEntity
    {
        public ButtonToRunEntity(string name, string toolTip)
        {
            Name = name;
            Tooltip = toolTip;
        }

        public string Name { get; private set; }

        public string Tooltip { get; private set; }
    }
}
