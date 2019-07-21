using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace binariex
{
    class Converter
    {
        const string EVAL_PREFIX = "eval:";
        readonly Dictionary<string, Action<XElement>> parseElemMap;

        readonly IReader reader;
        readonly IWriter writer;
        readonly XDocument schemaDoc;
        readonly V8ScriptEngine v8Engine;

        readonly Encoding encoding;
        readonly string endian;

        public Converter(IReader reader, IWriter writer, XDocument schemaDoc)
        {
            this.parseElemMap = new Dictionary<string, Action<XElement>>
            {
                { "if", ParseIf },
                { "sheet", ParseSheet },
                { "group", ParseGroup },
                { "leaf", ParseLeaf },
            };

            this.reader = reader;
            this.writer = writer;
            this.schemaDoc = schemaDoc;
            this.v8Engine = new V8ScriptEngine();

            this.encoding = Encoding.GetEncoding(schemaDoc.Root.Attribute("encoding")?.Value ?? "UTF-8");
            this.endian = schemaDoc.Root.Attribute("endian")?.Value ?? (BitConverter.IsLittleEndian ? "LE" : "BE");
        }

        public void Run()
        {
            ParseChildren(schemaDoc.Root.Element("data"));
        }
        public void Dispose()
        {
            this.v8Engine.Dispose();
        }

        void ParseElement(XElement elem)
        {
            var numRepeat = EvaluateExpr(elem.Attribute("repeat")?.Value) as int? ?? 1;
            var indexLabel = elem.Attribute("indexLabel")?.Value;
            bool flat = numRepeat == 1 || IsJSTrue(EvaluateExpr(elem.Attribute("flat")?.Value));
            if (!flat)
            {
                this.reader.BeginRepeat();
                this.writer.BeginRepeat();
            }
            for (int i = 1; i <= numRepeat; i++)
            {
                if (indexLabel != null)
                {
                    this.v8Engine.Script[indexLabel] = i;
                }

                if (elem.Attribute("if") != null && !IsJSTrue(EvaluateJSExpr(elem.Attribute("if").Value)))
                {
                    return;
                }

                if (this.parseElemMap.TryGetValue(elem.Name.LocalName, out var parseElem))
                {
                    parseElem(elem);
                }
                else
                {
                    throw new BinSchemaException(elem);
                }
            }
            if (!flat)
            {
                this.reader.EndRepeat();
                this.writer.EndRepeat();
            }
        }

        void ParseChildren(XElement elem)
        {
            foreach (var child in elem.Elements())
            {
                ParseElement(child);
            }
        }

        void ParseIf(XElement elem)
        {
            if (IsJSTrue(EvaluateJSExpr(elem.Attribute("cond").Value)))
            {
                ParseChildren(elem);
            }
        }

        void ParseSheet(XElement elem)
        {
            var name = elem.Attribute("name").Value;
            this.reader.PushSheet(name);
            this.writer.PushSheet(name);
            ParseChildren(elem);
            this.reader.PopSheet();
            this.writer.PopSheet();
        }

        void ParseGroup(XElement elem)
        {
            var name = elem.Attribute("name").Value;
            this.reader.PushGroup(name);
            this.writer.PushGroup(name);
            ParseChildren(elem);
            this.reader.PopGroup();
            this.writer.PopGroup();
        }

        void ParseLeaf(XElement elem)
        {
            var leafInfo = new LeafInfo {
                Name = EvaluateExpr(elem.Attribute("name").Value) as string,
                Type = EvaluateExpr(elem.Attribute("type").Value) as string,
                Size = EvaluateExpr(elem.Attribute("size").Value) as int? ?? 0,
                Encoding = elem.Attribute("encoding") != null ? Encoding.GetEncoding(elem.Attribute("encoding").Value) : this.encoding,
                Endian = EvaluateExpr(elem.Attribute("endian")?.Value) as string ?? this.endian,
            };
            if (leafInfo.Size <= 0)
            {
                throw new BinSchemaException(elem);
            }

            this.reader.GetValue(leafInfo, out var raw, out var decoded);
            this.writer.SetValue(leafInfo, raw, decoded);

            var label = elem.Attribute("label")?.Value;
            if (label != null)
            {
                this.v8Engine.Script[label] = decoded;
            }
        }

        object EvaluateExpr(string expr)
        {
            return expr == null ? null :
                expr.StartsWith(EVAL_PREFIX) ? EvaluateJSExpr(expr.Substring(EVAL_PREFIX.Length)) :
                Int32.TryParse(expr, out var outInt32) ? outInt32 as object :
                expr;
        }

        object EvaluateJSExpr(string expr)
        {
            return this.v8Engine.Evaluate(expr);
        }

        bool IsJSTrue(object value)
        {
            return !(
                value == null ||
                value is Undefined ||
                value is VoidResult ||
                value.Equals("") ||
                value.Equals(false) ||
                value.Equals(0)
            );
        }
    }
}
