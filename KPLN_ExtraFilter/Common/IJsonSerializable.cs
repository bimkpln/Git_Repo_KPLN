﻿namespace KPLN_ExtraFilter.Common
{
    public interface IJsonSerializable
    {
        /// <summary>
        /// Отдельный метод для чистки от лишних полей в классе при сериализации в Json
        /// </summary>
        object ToJson();
    }
}
