namespace binariex
{
    interface IReader
    {
        void BeginRepeat();
        void EndRepeat();
        void PopGroup();
        void PopSheet();
        void PushGroup(string name);
        void PushSheet(string name);
        void GetValue(LeafInfo leafInfo, out object raw, out object decoded);
        void Seek(long offset);
    }
}
