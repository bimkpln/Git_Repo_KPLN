using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Publication.Common
{
    public class ComboBoxSet
    {
        public bool Sheets { get; private set; }
        public bool Views { get; private set; }
        public bool IsUserCreated
        {
            get
            {
                if (!Sheets && !Views) { return true; }
                return false;
            }
        }
        public string Name { get; private set; }
        private Document Document { get; set; }
        public ObservableCollection<ListBoxElement> Elements
        {
            get
            {
                ObservableCollection<ListBoxElement> sortedElements = new ObservableCollection<ListBoxElement>();
                List<ListBoxElement> elements = new List<ListBoxElement>();
                if (Set != null)
                {
                    foreach (View i in Set.Views)
                    {
                        elements.Add(new ListBoxElement(i, true));
                    }
                }

                if (Views)
                {
                    foreach (View i in new FilteredElementCollector(Document).OfClass(typeof(View)).WhereElementIsNotElementType().ToElements())
                    {
                        if (i.CanBePrinted && i.GetType() != typeof(ViewSheet))
                        {
                            elements.Add(new ListBoxElement(i, false));
                        }
                    }
                }
                if (Sheets)
                {
                    foreach (ViewSheet i in new FilteredElementCollector(Document).OfClass(typeof(ViewSheet)).WhereElementIsNotElementType().ToElements())
                    {
                        elements.Add(new ListBoxElement(i, false));
                    }
                }
                elements.OrderBy(x => x.Name).ToList();
                foreach (ListBoxElement i in elements)
                { sortedElements.Add(i); }
                return sortedElements;
            }
        }
        public ViewSheetSet Set { get; private set; }
        public ComboBoxSet(Document doc, ViewSheetSet set)
        {
            //Print("ComboBoxSet", KPLN_Loader.Preferences.MessageType.System_OK);
            Document = doc;
            Sheets = false;
            Views = false;
            Name = set.Name;
            Set = set;
        }
        public ComboBoxSet(Document doc, string name, bool sheets, bool views)
        {
            //Print("ComboBoxSet", KPLN_Loader.Preferences.MessageType.System_OK);
            Document = doc;
            Sheets = sheets;
            Views = views;
            Name = name;
            Set = null;
        }
    }
}
