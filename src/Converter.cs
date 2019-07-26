using Jint;
using Jint.Native;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace binariex
{
    class Converter : IDisposable
    {
        static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        const string EVAL_PREFIX = "eval:";
        readonly Dictionary<string, Action<XElement>> parseElemMap;

        readonly IReader reader;
        readonly IWriter writer;
        readonly XDocument schemaDoc;

        readonly bool forward;

        readonly Encoding encoding;
        readonly string endian;

        readonly Dictionary<string, UserCode> codeMap = new Dictionary<string, UserCode>();

        readonly Engine jsEngine = new Engine();

        public Converter(IReader reader, IWriter writer, XDocument schemaDoc)
        {
            this.parseElemMap = new Dictionary<string, Action<XElement>>
            {
                { "if", ParseIf },
                { "seek", ParseSeek },
                { "sheet", ParseSheet },
                { "group", ParseGroup },
                { "leaf", ParseLeaf },
            };

            this.reader = reader;
            this.writer = writer;
            this.schemaDoc = schemaDoc;

            this.forward = reader is BinaryReader;

            try
            {
                this.encoding = GetEncoding(schemaDoc.Root.Attribute("encoding")?.Value ?? "UTF-8");
                this.endian = GetEndian(schemaDoc.Root.Attribute("endian")?.Value ?? (BitConverter.IsLittleEndian ? "LE" : "BE"));
            }
            catch (BinariexException exc)
            {
                throw exc.AddSchemaElement(schemaDoc.Root);
            }
        }

        public void Run()
        {
            if (schemaDoc.Root.Element("defs") != null)
            {
                ParseDefs(schemaDoc.Root.Element("defs"));
            }
            ParseChildren(schemaDoc.Root.Element("data"));
        }

        public void Dispose()
        {
            // Nothing to do
        }

        void ParseDefs(XElement elem)
        {
            try
            {
                foreach (var scriptElem in elem.Elements("script"))
                {
                    this.jsEngine.Execute(scriptElem.Value);
                }
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "reading script");
            }

            foreach (var codeElem in elem.Elements("code"))
            {
                var name = GetAttr(codeElem, "name");
                var codeObj = UserCode.Create(codeElem, this.forward, this.jsEngine);
                this.codeMap.Add(name, codeObj);
            }
        }

        void ParseElementSafe(XElement elem)
        {
            try
            {
                ParseElement(elem);
            }
            catch (BinariexException exc)
            {
                throw exc.SchemaLineInfo == null ? exc.AddSchemaElement(elem) : exc;
            }
        }

        void ParseElement(XElement elem)
        {
            var numRepeatRepr = elem.Attribute("repeat")?.Value;
            var numRepeat = numRepeatRepr == "*" ? Int32.MaxValue :
                EvaluateExpr(numRepeatRepr) as int? ?? 1;
            var indexLabel = elem.Attribute("indexLabel")?.Value;
            bool flat = numRepeat == 1 || IsJSTrue(EvaluateExpr(elem.Attribute("flat")?.Value));
            if (!flat)
            {
                this.reader.BeginRepeat();
                this.writer.BeginRepeat();
            }
            for (int i = 0; i < numRepeat; i++)
            {
                if (elem.Attribute("if") != null && !IsJSTrue(EvaluateJSExpr(elem.Attribute("if").Value)))
                {
                    return;
                }

                if (indexLabel != null)
                {
                    this.jsEngine.SetValue(indexLabel, i);
                }

                if (!this.parseElemMap.TryGetValue(elem.Name.LocalName, out var parseElem))
                {
                    throw new BinariexException("reading schema", "Invalid schema element: {0}.", elem.Name.LocalName);
                }

                try
                {
                    parseElem(elem);
                }
                catch (EndOfStreamException)
                {
                    if (numRepeatRepr == "*")
                    {
                        break;
                    }
                    else
                    {
                        throw;
                    }
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
                ParseElementSafe(child);
            }
        }

        void ParseIf(XElement elem)
        {
            if (IsJSTrue(EvaluateJSExpr(GetAttr(elem, "cond"))))
            {
                ParseChildren(elem);
            }
        }

        void ParseSeek(XElement elem)
        {
            var offsetRepr = GetAttr(elem, "offset");
            if (!long.TryParse(offsetRepr, out var offset))
            {
                throw new BinariexException("reading schema", "Invalid number format: {0}", offsetRepr);
            }
            try
            {
                this.reader.Seek(offset);
                this.writer.Seek(offset);
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "seeking");
            }
        }

        void ParseSheet(XElement elem)
        {
            var name = GetAttr(elem, "name");
            this.reader.PushSheet(name);
            this.writer.PushSheet(name);
            ParseChildren(elem);
            this.reader.PopSheet();
            this.writer.PopSheet();
        }

        void ParseGroup(XElement elem)
        {
            var name = GetAttr(elem, "name");
            this.reader.PushGroup(name);
            this.writer.PushGroup(name);
            ParseChildren(elem);
            this.reader.PopGroup();
            this.writer.PopGroup();
        }

        void ParseLeaf(XElement elem)
        {
            var leafInfo = new LeafInfo {
                Name = EvaluateExpr(GetAttr(elem, "name")) as string,
                Type = EvaluateExpr(GetAttr(elem, "type")) as string,
                Size = EvaluateExpr(GetAttr(elem, "size")) as int? ?? 0,
                Encoding = elem.Attribute("encoding") != null ? GetEncoding(elem.Attribute("encoding").Value) : this.encoding,
                Endian = elem.Attribute("endian") != null ? GetEndian(EvaluateExpr(elem.Attribute("endian").Value) as string) : this.endian,
                HasUserCode = elem.Attribute("code") != null
            };
            if (leafInfo.Size <= 0)
            {
                throw new BinariexException("reading schema", "Invalid size: {0}", leafInfo.Size);
            }

            GetValueSafe(leafInfo, out var raw, out var firstObj, out var output);

            var secondObj = null as object;
            if (output == null)
            {
                var codeName = elem.Attribute("code")?.Value;
                secondObj = GetSecondObject(firstObj, codeName);
            }

            SetValueSafe(leafInfo, raw, secondObj, output);

            var label = elem.Attribute("label")?.Value;
            if (label != null)
            {
                this.jsEngine.SetValue(label, raw is byte[] ? secondObj : firstObj);
            }
        }

        void GetValueSafe(LeafInfo leafInfo, out object raw, out object firstObj, out object output)
        {
            try
            {
                this.reader.GetValue(leafInfo, out raw, out firstObj, out output);
            }
            catch (EndOfStreamException)
            {
                throw;
            }
            catch (BinariexException)
            {
                throw;
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "reading input file");
            }
        }

        object GetSecondObject(object firstObj, string codeName)
        {
            if (codeName == null)
            {
                return firstObj;
            }

            if (!this.codeMap.TryGetValue(codeName, out var codeObj))
            {
                throw new BinariexException("reading schema", "Invalid code name: {0}", codeName);
            }

            var secondObj = codeObj.Convert(firstObj);
            if (secondObj == null)
            {
                throw new BinariexException(
                    "reading input file", "Invalid value for code {0}: {1}",
                    codeName,
                    firstObj is byte[]? string.Join("", (firstObj as byte[]).Select(e => e.ToString("X2"))) :
                    firstObj is string && Regex.IsMatch(firstObj as string, "^\\s*$") ? $@"'{firstObj}'" : firstObj
                );
            }

            return secondObj;
        }

        void SetValueSafe(LeafInfo leafInfo, object raw, object secondObj, object output)
        {
            try
            {
                this.writer.SetValue(leafInfo, raw, secondObj, output);
            }
            catch (BinariexException)
            {
                throw;
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "writing to file");
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
            try
            {
                return this.jsEngine.Execute(expr).GetCompletionValue().ToObject();
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "evaluating script");
            }
        }

        bool IsJSTrue(object value)
        {
            return !(
                value == null ||
                value.Equals("") ||
                value.Equals(false) ||
                value.Equals(0)
            );
        }

        static string GetAttr(XElement elem, string attrName)
        {
            return elem.Attribute(attrName)?.Value ??
                throw new BinariexException("reading schema", "Attribute '{0}' not found.", attrName).AddSchemaElement(elem);
        }

        static Encoding GetEncoding(string encodingName)
        {
            try
            {
                return Encoding.GetEncoding(encodingName);
            }
            catch (ArgumentException exc)
            {
                throw new BinariexException(exc, "reading schema");
            }
        }

        static string GetEndian(string endianName)
        {
            if (!(endianName == "LE" || endianName == "BE"))
            {
                throw new BinariexException("reading schema", "Endian must be 'LE' or 'BE'.");
            }
            return endianName;
        }

        class UserCode
        {
            readonly Dictionary<object, object> codeMap;
            readonly object defaultValue;

            readonly Engine jsEngine;
            readonly JsValue codeScript;

            public static UserCode Create(XElement codeElem, bool forward, Engine jsEngine)
            {
                var encodeType = GetAttr(codeElem, "type");
                var codeMap = new Dictionary<object, object>();
                foreach (XElement valElem in codeElem.Elements("codeval"))
                {
                    var encoded = ParseEncodedRepr(GetAttr(valElem, "encoded"), encodeType, forward);
                    var decoded = GetAttr(valElem, "decoded");

                    var key = forward ? encoded : decoded;
                    var value = forward ? decoded : encoded;
                    if (!codeMap.ContainsKey(key))
                    {
                        codeMap.Add(key, value);
                    }
                }

                var codeScript = null as JsValue;
                var scriptElem = codeElem.Element(forward ? "decode" : "encode");
                if (scriptElem != null)
                {
                    try
                    {
                        codeScript = jsEngine.Execute($@"(function($input) {{{scriptElem.Value}}})").GetCompletionValue();
                    }
                    catch (Exception exc)
                    {
                        throw new BinariexException(exc, "reading code definition").AddSchemaElement(scriptElem);
                    }
                }

                var defaultValue = forward ?
                    codeElem.Element("default")?.Attribute("decoded")?.Value :
                    ParseEncodedRepr(codeElem.Element("default")?.Attribute("encoded")?.Value, encodeType, forward);

                return new UserCode(codeMap, defaultValue, jsEngine, codeScript);
            }

            UserCode(Dictionary<object, object> codeMap, object defaultValue, Engine jsEngine, JsValue codeScript)
            {
                this.codeMap = codeMap;
                this.defaultValue = defaultValue;
                this.jsEngine = jsEngine;
                this.codeScript = codeScript;
            }

            public object Convert(object input)
            {
                if (this.codeMap.Count > 0)
                {
                    var inputRepr = input is byte[] ? string.Join("", (input as byte[]).Select(e => e.ToString("X2")).ToArray()) : input;
                    if (this.codeMap.TryGetValue(inputRepr, out var converted))
                    {
                        return converted;
                    }
                }

                if (this.codeScript != null)
                {
                    try
                    {
                        var inputJsValue = JsValue.FromObject(jsEngine, input);
                        var converted = this.codeScript.Invoke(inputJsValue);
                        if (!(converted.IsUndefined()))
                        {
                            return converted.ToObject();
                        }
                    }
                    catch (Exception exc)
                    {
                        throw new BinariexException(exc, "evaluating code script");
                    }
                }

                return this.defaultValue;
            }

            static object ParseEncodedRepr(string encodedRepr, string encodeType, bool forward)
            {
                return
                    encodedRepr == null ? null :
                    encodeType == "bin" && !forward ? Enumerable.Range(0, encodedRepr.Length / 2).Select(i => System.Convert.ToByte(encodedRepr.Substring(i * 2, 2), 16)).ToArray() :
                    encodeType == "char" ? Regex.Replace(encodedRepr, @"\\x(\d{1,4})", m => char.ConvertFromUtf32(System.Convert.ToInt32(m.Groups[1].Value, 16))) :
                    encodeType == "int" ? long.Parse(encodedRepr) as object :
                    encodeType == "uint" ? ulong.Parse(encodedRepr) as object :
                    encodedRepr;
            }
        }
    }
}
