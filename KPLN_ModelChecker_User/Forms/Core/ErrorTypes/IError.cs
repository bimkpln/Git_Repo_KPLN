namespace KPLN_ModelChecker_User.Forms.Core.ErrorTypes
{
    public interface IError
    {
        /// <summary>
        /// Идентификатор критичности ошибки. Начиная с 1
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Имя ошибки
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Описание ошибки
        /// </summary>
        string Description { get; }
    }
}
