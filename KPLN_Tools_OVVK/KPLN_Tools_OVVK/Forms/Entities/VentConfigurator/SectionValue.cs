namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class SectionValue : EditableValue
    {
        private SectionTypeChoice _type;
        internal bool IsConnector, IsValve;
        internal DimensionValue Width, Height, OffsetX, OffsetY;
        internal DimensionValue Length;

        internal bool IsFixed
        {
            get
            {
                return IsConnector || IsValve;
            }
        }

        public SectionTypeChoice Type
        {
            get
            {
                return _type;
            }
            set
            {
                if (ReferenceEquals(_type, value))
                    return;
                bool changed = _type?.Key != value?.Key;
                _type = value;
                if (changed)
                {
                    ResetEmptyLength();
                    Changed();
                }
            }
        }

        internal void ResetEmptyLength()
        {
            // Начальное значение при выборе пустого блока. Дальше длина редактируется как обычно.
            if (Type != null && Type.IsEmpty && Length != null)
                Length.Text = "1";
        }

        internal SectionValue Copy()
        {
            return new SectionValue
            {
                _type = _type,
                IsConnector = IsConnector,
                IsValve = IsValve,
                Length = Length.Copy(),
                Width = Width?.Copy(),
                Height = Height?.Copy(),
                OffsetX = OffsetX?.Copy(),
                OffsetY = OffsetY?.Copy()
            };
        }
    }
}
