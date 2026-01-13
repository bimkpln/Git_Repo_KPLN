namespace KPLN_Loader.Services
{
    /// <summary>
    /// Структура для описания таблиц, используемых в MainDB
    /// </summary>
    public enum MainDB_Tables
    {
        Documents,
        LoaderDescriptions,
        Modules,
        Projects,
        ProjectsMatrix,
        SubDepartments,
        Users
    }

    /// <summary>
    /// Варианты реакций на сообщения
    /// </summary>
    public enum MainDB_LoaderDescriptions_RateType
    {
        Approval,
        Disapproval,
    }
}
