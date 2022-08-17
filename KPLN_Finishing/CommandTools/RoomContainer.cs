using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Finishing.Tools;

namespace KPLN_Finishing.CommandTools
{
    class RoomContainer
    {
        public static List<AbstractFinishingFilter> Filters = new List<AbstractFinishingFilter>();
        public static bool AddTypeKeys = true;
        public readonly List<Element> Walls = new List<Element>();
        public readonly List<Element> Floors = new List<Element>();
        public readonly List<Element> Ceilings = new List<Element>();
        private List<string> Ids = new List<string>();
        public string Keys
        {
            get
            {
                return string.Join("_", Ids);
            }
        }
        public Room Room { get; }
        public RoomContainer(Room room)
        {
            Room = room;
        }
        public void InsertItem(Element item)
        {
            switch (item.Category.Id.IntegerValue)
            {
                case -2000011:
                    Walls.Add(item);
                    break;
                case -2000032:
                    Floors.Add(item);
                    break;
                case -2000038:
                    Ceilings.Add(item);
                    break;
                default:
                    return;
            }
            string id = GetTypeElement(item).Id.IntegerValue.ToString();
            if (!Ids.Contains(id)) { Ids.Add(id); }
            Ids.Sort();
        }
    }
}
