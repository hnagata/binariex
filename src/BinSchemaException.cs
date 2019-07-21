using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace binariex
{
    class BinSchemaException : Exception
    {
        public XElement TargetElement { get; private set; }

        public BinSchemaException(XElement elem)
        {
            this.TargetElement = elem;
        }

        BinSchemaException(string message, XElement elem) : base(message)
        {
            this.TargetElement = elem;
        }

        BinSchemaException(string message, Exception inner, XElement elem): base(message, inner)
        {
            this.TargetElement = elem;
        }
    }
}
