using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.Core
{
    /// /// <summary>
    /// Сущность эл-в АР и КР для кэширования информации и объединения IOSElemEntity по эл-ту АР/КР
    /// </summary>
    internal sealed class ARKRElemEntity : InstanceGeomData
    {
        internal ARKRElemEntity(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Ссылка на элементы IOSElemEntity для сущности
        /// </summary>
        internal List<IOSElemEntity> IOSElemEntities { get; private set; } = new List<IOSElemEntity>();
    }
}
