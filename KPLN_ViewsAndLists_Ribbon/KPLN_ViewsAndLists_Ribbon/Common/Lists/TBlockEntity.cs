using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_ViewsAndLists_Ribbon.Common.Lists
{
    /// <summary>
    /// Сущность для плагина CommandListTBlockParamCopier
    /// </summary>
    internal class TBlockEntity
    {
        public TBlockEntity(ViewSheet currentViewSheet, IEnumerable<Element> currentTBlocks)
        {
            CurrentViewSheet = currentViewSheet;
            CurrentTBlocks = currentTBlocks;
        }
        
        internal ViewSheet CurrentViewSheet { get; }
        
        internal IEnumerable<Element> CurrentTBlocks { get; }

    }
}
