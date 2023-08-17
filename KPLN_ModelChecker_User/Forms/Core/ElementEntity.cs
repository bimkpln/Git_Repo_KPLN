using Autodesk.Revit.DB;
using KPLN_ModelChecker_User.Forms.Core.ErrorTypes;

namespace KPLN_ModelChecker_User.Forms.Core
{
    /// <summary>
    /// Общая сущность для генерации элементов с ошибками
    /// </summary>
    public sealed class ElementEntity
    {
        /// <summary>
        /// Элемент из Ревит
        /// </summary>
        public Element Element { get; private set; }

        /// <summary>
        /// Имя элемента Ревит
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Название ошибки 
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Описание ошибки
        /// </summary>
        public string Description { get; private set; }

        /// <summary>
        /// Тип ошибки
        /// </summary>
        public IError ErrorStatus { get; private set; }

        /// <summary>
        /// Инструкция по устранению ошибки
        /// </summary>
        public string Action { get; private set; }

        /// <summary>
        /// Общая сущность для генерации элементов с ошибками
        /// </summary>
        /// <param name="element">Элемент ревит</param>
        /// <param name="name">Имя элемента Ревит</param>
        /// <param name="title">Название ошибки</param>
        /// <param name="description">Описание ошибки</param>
        /// <param name="error">Тип ошибки</param>
        public ElementEntity(Element element, string name, string title, string description, IError error)
        {
            Element = element;
            Name = name;
            Title = title;
            Description = description;
            ErrorStatus = error;
        }

        /// <summary>
        /// Общая сущность для генерации элементов с ошибками
        /// </summary>
        /// <param name="element">Элемент ревит</param>
        /// <param name="name">Имя элемента Ревит</param>
        /// <param name="title">Название ошибки</param>
        /// <param name="description">Описание ошибки</param>
        /// <param name="error">Тип ошибки</param>
        /// <param name="action">Инструкция по устранению ошибки</param>
        public ElementEntity(Element element, string name, string title, string description, IError error, string action) : this(element, name, title, description, error)
        {
            Action = action;
        }
    }
}
