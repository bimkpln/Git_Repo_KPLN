namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class FileEntity
    {
        public FileEntity(string name, string path)
        {
            Name = name;
            Path = path;
        }

        public string Name { get; set; }

        public string Path { get; set; }
    }
}
