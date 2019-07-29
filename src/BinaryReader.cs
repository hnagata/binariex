using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace binariex
{
    class BinaryReader : IReader, IDisposable
    {
        readonly Dictionary<string, Func<LeafInfo, byte[], object>> decodeMap;
        readonly FileStream stream;

        public BinaryReader(string path)
        {
            this.decodeMap = new Dictionary<string, Func<LeafInfo, byte[], object>>
            {
                { "bin", DecodeBinary },
                { "char", DecodeString },
                { "int", DecodeSignedInteger },
                { "uint", DecodeUnsignedInteger },
                { "pbcd", DecodePBCD }
            };

            this.stream = File.OpenRead(path);
        }

        public void BeginRepeat()
        {
            // Nothing to do
        }

        public void EndRepeat()
        {
            // Nothing to do
        }

        public void PushGroup(string name)
        {
            // Nothing to do
        }

        public void PushSheet(string name)
        {
            // Nothing to do
        }

        public void PopGroup()
        {
            // Nothing to do
        }

        public void PopSheet()
        {
            // Nothing to do
        }

        public void GetValue(LeafInfo leafInfo, out object raw, out object decoded, out object output)
        {
            var buffer = new byte[leafInfo.Size];
            var total = 0;
            var read = -1;
            while (read != 0 && total < leafInfo.Size)
            {
                read = this.stream.Read(buffer, total, leafInfo.Size - total);
                total += read;
            }
            if (total == 0)
            {
                throw new EndOfStreamException();
            } else if (total < leafInfo.Size)
            {
                throw new BinariexException("reading input file", "Reached EOF before schema end.");
            }
            raw = buffer;

            if (!this.decodeMap.TryGetValue(leafInfo.Type, out var decode))
            {
                throw new BinariexException("reading schema", "Invalid data type: {0}", leafInfo.Type);
            }

            decoded = decode(leafInfo, buffer);
            output = null;
        }

        public void Seek(long offset)
        {
            this.stream.Seek(offset, SeekOrigin.Current);
        }

        public long GetReadPosition()
        {
            return this.stream.Position;
        }

        public long GetTotalSize()
        {
            return this.stream.Length;
        }

        public void Dispose()
        {
            this.stream.Dispose();
        }

        object DecodeBinary(LeafInfo leafInfo, byte[] raw)
        {
            return raw;
        }

        object DecodeString(LeafInfo leafInfo, byte[] raw)
        {
            return leafInfo.Encoding.GetString(raw);
        }

        object DecodeSignedInteger(LeafInfo leafInfo, byte[] raw)
        {
            try
            {
                var ordered = (BitConverter.IsLittleEndian ^ leafInfo.Endian == "LE") ? raw.Reverse().ToArray() : raw;
                var x =
                    raw.Length == 1 ? (raw[0] & 0x80) == 0 ? raw[0] & 0x7f : (raw[0] & 0x7f) - 0x80 :
                    raw.Length == 2 ? BitConverter.ToInt16(ordered, 0) :
                    raw.Length == 4 ? BitConverter.ToInt32(ordered, 0) :
                    raw.Length == 8 ? BitConverter.ToInt64(ordered, 0) :
                    throw new BinariexException("reading schema", "Invalid data length: {0}", raw.Length);
                return x;
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "Invalid data value");
            }
        }

        object DecodeUnsignedInteger(LeafInfo leafInfo, byte[] raw)
        {
            try
            {
                var ordered = (BitConverter.IsLittleEndian ^ leafInfo.Endian == "LE") ? raw.Reverse().ToArray() : raw;
                var x =
                    raw.Length == 1 ? raw[0] :
                    raw.Length == 2 ? BitConverter.ToUInt16(ordered, 0) :
                    raw.Length == 4 ? BitConverter.ToUInt32(ordered, 0) :
                    raw.Length == 8 ? BitConverter.ToUInt64(ordered, 0) :
                    throw new BinariexException("reading schema", "Invalid data length: {0}", raw.Length);
                return x;
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "Invalid data value");
            }
        }

        object DecodePBCD(LeafInfo leafInfo, byte[] raw)
        {
            long x = 0;
            for (var i = 0; i < raw.Length - 1; i++)
            {
                x = x * 100 + (raw[i] >> 4 & 0x0f) * 10 + (raw[i] & 0x0f);
            }
            x = x * 10 + (raw[raw.Length - 1] >> 4 & 0x0f);

            var signCode = raw[raw.Length - 1] & 0x0f;
            if (signCode == 0x0b || signCode == 0x0d)
            {
                x = -x;
            }

            return x;
        }
    }
}
