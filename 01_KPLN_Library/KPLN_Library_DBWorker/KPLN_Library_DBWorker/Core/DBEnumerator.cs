namespace KPLN_Library_DBWorker.Core
{
    /// <summary>
    /// Коллекция используемых баз данных
    /// </summary>
    public enum DBEnumerator
    {
        Documents,
        LoaderDescriptions,
        LoaderDescriptions_Single,
        ModuleAutostart,
        Modules,
        Projects = 6, // Сохраняем значения enum после удаления устаревшего пункта.
        ProjectsAccessMatrix,
        ProjectsIOSClashMatrix,
        RevitDialogs,
        RevitDocExchanges,
        SubDepartmentCodeMatrix,
        SubDepartments,
        Users
    }
}
