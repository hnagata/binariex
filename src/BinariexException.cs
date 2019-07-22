using System;
using System.Xml;
using System.Xml.Linq;

namespace binariex
{
    class BinariexException : Exception
    {
        public string Stage { get; private set; }
        public object[] MessageParams { get; private set; }
        public string InputPath { get; private set; }
        public string SchemaPath { get; private set; }
        public IXmlLineInfo SchemaLineInfo { get; private set; }
        public string ReaderPosition { get; private set; }
        public string WriterPosition { get; private set; }

        public BinariexException(Exception inner, string stage) : base(inner.Message, inner)
        {
            this.Stage = stage;
        }

        public BinariexException(string stage, string message, params string[] messageParams) : base(message)
        {
            this.Stage = stage;
            this.MessageParams = messageParams;
        }

        public BinariexException(Exception inner, string stage, string message, params string[] messageParams) : base(message, inner)
        {
            this.Stage = stage;
            this.MessageParams = messageParams;
        }

        public BinariexException AddInputPath(string path)
        {
            this.InputPath = path;
            return this;
        }

        public BinariexException AddSchemaPath(string path)
        {
            this.SchemaPath = path;
            return this;
        }

        public BinariexException AddSchemaElement(XElement elem)
        {
            this.SchemaLineInfo = elem as IXmlLineInfo;
            return this;
        }

        public BinariexException AddReaderPosition(string pos)
        {
            this.ReaderPosition = pos;
            return this;
        }

        public BinariexException AddWriterPosition(string pos)
        {
            this.WriterPosition = pos;
            return this;
        }
    }
}
