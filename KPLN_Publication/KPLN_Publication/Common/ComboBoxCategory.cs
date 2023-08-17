namespace KPLN_Publication.Common
{
    public class ComboBoxCategory
    {
        public bool Sheets { get; private set; }
        public bool Views { get; private set; }
        public string Name { get; private set; }
        public ComboBoxCategory(string name, bool sheets, bool views)
        {
            //Print("ComboBoxCategory", KPLN_Loader.Preferences.MessageType.System_OK);
            Name = name;
            Sheets = sheets;
            Views = views;
        }
    }
}
