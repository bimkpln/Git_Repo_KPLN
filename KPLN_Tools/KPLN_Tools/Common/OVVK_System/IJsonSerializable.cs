namespace KPLN_Tools.Common.OVVK_System
{
    public interface IJsonSerializable
    {
        /// <summary>
        /// Отдельный метод для чистки от лишних полей в классе при сериализации в Json
        /// </summary>
        object ToJson();
    }
}
