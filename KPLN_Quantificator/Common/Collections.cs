using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Quantificator.Common
{
    public static class Collections
    {
        public enum GroupingMode
        {
            [Description("<None>")]
            None,
            [Description("Уровень")]
            Level,
            [Description("Grid Intersection")]
            GridIntersection,
            [Description("Выбранные A")]
            SelectionA,
            [Description("Выбранные B")]
            SelectionB,
            [Description("Предназначено кому")]
            AssignedTo,
            [Description("Одобрено кем")]
            ApprovedBy,
            [Description("Статус")]
            Status,
            [Description("Основа")]
            Host
        }
    }
}
