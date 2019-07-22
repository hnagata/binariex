using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace binariex
{
    class BinaryWriter : IWriter, IDisposable
    {
        readonly Dictionary<string, Func<LeafInfo, object, byte[]>> encodeMap;

        readonly FileStream stream;

        public BinaryWriter(string path)
        {
            this.encodeMap = new Dictionary<string, Func<LeafInfo, object, byte[]>>
            {
                { "bin", EncodeBinary },
                { "char", EncodeString },
                { "int", EncodeSignedInteger },
                { "uint", EncodeUnsignedInteger },
                { "pbcd", EncodePBCD }
            };

            this.stream = File.OpenWrite(path);
        }

        public void BeginRepeat()
        {
            // Nothing to do
        }

        public void EndRepeat()
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

        public void PushGroup(string name)
        {
            // Nothing to do
        }

        public void PushSheet(string name)
        {
            // Nothing to do
        }

        public void SetValue(LeafInfo leafInfo, object raw, object decoded)
        {
            if (!this.encodeMap.TryGetValue(leafInfo.Type, out var encode))
            {
                throw new InvalidDataException();
            }

            var encoded = encode(leafInfo, decoded);
            if (encoded.Length != leafInfo.Size)
            {
                throw new InvalidDataException();
            }

            this.stream.Write(encoded, 0, encoded.Length);
        }

        public void Seek(long offset)
        {
            this.stream.Seek(offset, SeekOrigin.Current);
        }

        public void Dispose()
        {
            this.stream.Dispose();
        }

        byte[] EncodeBinary(LeafInfo leafInfo, object decoded)
        {
            return decoded as byte[];
        }

        byte[] EncodeString(LeafInfo leafInfo, object decoded)
        {
            return leafInfo.Encoding.GetBytes(decoded as string);
        }

        byte[] EncodeSignedInteger(LeafInfo leafInfo, object decoded)
        {
            var decodedInt = decoded as Int64?;
            var buffer =
                leafInfo.Size == 1 ? new byte[] { (byte)decodedInt } :
                leafInfo.Size == 2 ? BitConverter.GetBytes((short)decodedInt) :
                leafInfo.Size == 4 ? BitConverter.GetBytes((int)decodedInt) :
                leafInfo.Size == 8 ? BitConverter.GetBytes((long)decodedInt) :
                throw new InvalidDataException();
            return (BitConverter.IsLittleEndian ^ leafInfo.Endian == "LE") ? buffer.Reverse().ToArray() : buffer;
        }

        byte[] EncodeUnsignedInteger(LeafInfo leafInfo, object decoded)
        {
            var decodedInt = decoded as UInt64?;
            var buffer =
                leafInfo.Size == 1 ? new byte[] { (byte)decodedInt } :
                leafInfo.Size == 2 ? BitConverter.GetBytes((ushort)decodedInt) :
                leafInfo.Size == 4 ? BitConverter.GetBytes((uint)decodedInt) :
                leafInfo.Size == 8 ? BitConverter.GetBytes((ulong)decodedInt) :
                throw new InvalidDataException();
            return (BitConverter.IsLittleEndian ^ leafInfo.Endian == "LE") ? buffer.Reverse().ToArray() : buffer;
        }

        byte[] EncodePBCD(LeafInfo leafInfo, object decoded)
        {
            var decodedInt = (long)decoded;
            byte[] buffer = new byte[leafInfo.Size];

            buffer[leafInfo.Size - 1] = (byte)(decodedInt >= 0 ? 0x0c : 0x0d);
            decodedInt = decodedInt >= 0 ? decodedInt : -decodedInt;

            buffer[leafInfo.Size - 1] |= (byte)((decodedInt % 10) << 4);
            decodedInt /= 10;

            for (var i = buffer.Length - 2; i >= 0 && decodedInt > 0; i--)
            {
                buffer[i] = (byte)(
                    (decodedInt % 10) |
                    (decodedInt / 10 % 10) << 4
                );
                decodedInt /= 100;
            }

            return buffer;
        }
    }
}
