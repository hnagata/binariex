namespace binariex
{
    interface IWriter
    {
        void BeginRepeat();
        void EndRepeat();
        void PopGroup();
        void PopSheet();
        void PushGroup(string name);
        void PushSheet(string name);
        void SetValue(LeafInfo leafInfo, object raw, object decoded);
    }
}
