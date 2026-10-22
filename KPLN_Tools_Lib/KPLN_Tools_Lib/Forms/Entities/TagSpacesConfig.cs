using KPLN_Library_ConfigWorker.Core;

namespace KPLN_Tools_Lib.Forms.Entities
{
    /// <summary>Только сохраняемые настройки, без объектов Revit и выбранных видов.</summary>
    public sealed class TagSpacesConfig : IJsonSerializable
    {
        public bool IsRoom { get; set; } = true;
        public TagSpacesTypeConfig RoomTag { get; set; }
        public TagSpacesTypeConfig SpaceTag { get; set; }
        public TagSpacesSourceConfig[] Sources { get; set; }
        public bool UseLinkedRooms { get; set; }
        public string LinkUniqueId { get; set; }
        public string LinkName { get; set; }
        public bool UseRadiators { get; set; }
        public bool UseGrilles { get; set; }
        public TagSpacesTypeConfig RadiatorTag { get; set; }
        public TagSpacesTypeConfig GrilleTag { get; set; }

        public object ToJson() => this;
    }

    public sealed class TagSpacesSourceConfig
    {
        public string LinkUniqueId { get; set; }
        public string LinkName { get; set; }
    }

    public sealed class TagSpacesTypeConfig
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
    }
}
