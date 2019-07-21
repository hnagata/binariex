using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

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
    }
}
