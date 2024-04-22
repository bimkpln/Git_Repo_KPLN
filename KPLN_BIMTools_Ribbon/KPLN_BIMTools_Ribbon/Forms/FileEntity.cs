namespace KPLN_BIMTools_Ribbon.Forms
{
    internal class FileEntity
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
