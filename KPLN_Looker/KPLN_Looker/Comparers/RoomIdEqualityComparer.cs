using Autodesk.Revit.DB.Architecture;
using System.Collections.Generic;

namespace KPLN_Looker.Comparers
{
    internal sealed class RoomIdEqualityComparer : IEqualityComparer<Room>
    {
        public static readonly RoomIdEqualityComparer Instance = new RoomIdEqualityComparer();
        private RoomIdEqualityComparer() { }

        public bool Equals(Room x, Room y)
        {
            if (ReferenceEquals(x, y)) 
                return true;
            
            if (x is null || y is null)
                return false;
            
            return ElementIdEqualityComparer.Instance.Equals(x.Id, y.Id);
        }

        public int GetHashCode(Room obj) => ElementIdEqualityComparer.Instance.GetHashCode(obj?.Id);
    }
}
