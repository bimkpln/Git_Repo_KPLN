using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ViewsAndLists_Ribbon.Common.Lists
{

    /// <summary>
    /// Коллекция кодов Юникодов Revit
    /// </summary>
    public enum UniDecCodes
    {
        RS = 30,
        US = 31,
        ZWJN = 8204,
        ZWJ = 8205,
        LRM = 8206,
        RLM = 8207,
        LRE = 8234,
        RLE = 8235,
        LRO = 8236,
        PDF = 8237,
        RLO = 8238,
        ISS = 8298,
        ASS = 8299,
        IAFS = 8301,
        AAFS = 8301,
        NADS = 8302,
        NODS = 8303
    }

    public sealed class UniEntity
    {
        public string Name { get; set; }
        public string Code { get; set; }
        public int DecCode { get; set; }
    }
}
