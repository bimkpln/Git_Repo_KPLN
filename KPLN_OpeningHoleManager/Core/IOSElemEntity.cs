using Autodesk.Revit.DB;

namespace KPLN_OpeningHoleManager.Core
{
    public class IOSElemEntity
    {
        public IOSElemEntity(Document linkDoc, Transform linkTrans, Element elem, Solid elemSolid, Solid aRIOS_IntesectionSolid)
        {
            IOS_LinkDocument = linkDoc;
            IOS_LinkTransform = linkTrans;
            IOS_Element = elem;
            IOS_Solid = elemSolid;
            ARIOS_IntesectionSolid = aRIOS_IntesectionSolid;
        }

        /// <summary>
        /// Ссылка на документ линка
        /// </summary>
        public Document IOS_LinkDocument { get; private set; }

        /// <summary>
        /// Ссылка на Transform для линка
        /// </summary>
        public Transform IOS_LinkTransform { get; private set; }

        /// <summary>
        /// Ссылка на элемент модели
        /// </summary>
        public Element IOS_Element { get; private set; }

        /// <summary>
        /// Кэширование SOLID геометрии
        /// </summary>
        public Solid IOS_Solid { get; private set; }

        /// <summary>
        /// Кэширование SOLID геометрии ПЕРЕСЕЧЕНИЯ между АР и ИОС
        /// </summary>
        public Solid ARIOS_IntesectionSolid { get; private set; }
    }
}
