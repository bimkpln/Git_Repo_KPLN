using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Services;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.Core
{
    /// /// <summary>
    /// Сущность эл-в АР и КР для кэширования информации и объединения IOSElemEntity по эл-ту АР/КР
    /// </summary>
    internal sealed class ARKRElemEntity
    {
        internal ARKRElemEntity(Element aRKR_Element)
        {
            ARKRHost_Element = aRKR_Element;
            ARKRHost_Solid = GeometryWorker.GetRevitElemSolid(aRKR_Element);
        }

        /// <summary>
        /// Ссылка на элемент АР/КР
        /// </summary>
        internal Element ARKRHost_Element { get; private set; }

        /// <summary>
        /// Кэширование SOLID геометрии АР/КР
        /// </summary>
        internal Solid ARKRHost_Solid { get; private set; }

        /// <summary>
        /// Ссылка на элементы IOSElemEntity для сущности
        /// </summary>
        internal List<IOSElemEntity> IOSElemEntities { get; private set; } = new List<IOSElemEntity>();
    }
}
