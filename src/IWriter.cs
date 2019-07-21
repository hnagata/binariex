using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

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
