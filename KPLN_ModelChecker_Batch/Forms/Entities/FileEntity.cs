namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class FileEntity
    {
        public FileEntity(string name, string path, string fileSize)
        {
            Name = name;
            Path = path;
            FileSize = fileSize;
        }

        /// <summary>
        /// Имя файла
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Путь к файлу
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Размер файла
        /// </summary>
        public string FileSize { get; set; }
    }
}
