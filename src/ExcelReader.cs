using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace binariex
{
    class ExcelReader : ExcelBase, IReader, IDisposable
    {
        string path;
        ExcelPackage package;

        public ExcelReader(string path)
        {
            this.path = path;
            this.package = new ExcelPackage(new FileInfo(path));
        }

        protected override ExcelWorksheet OpenSheet(string name)
        {
            return package.Workbook.Worksheets[name];
        }

        public void GetValue(LeafInfo leafInfo, out object raw, out object decoded, out object output)
        {
            var ctx = this.CurrentContext;
            if (ctx == null)
            {
                throw new InvalidOperationException();
            }

            if (ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex > ctx.Sheet.Dimension.Rows)
            {
                throw new EndOfStreamException();
            }

            var cell = ctx.Sheet.Cells[ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex, ctx.CursorColumnIndex];
            var cellText = cell.Text;
            raw = cellText;

            try
            {
                GetDecodedValue(leafInfo, cellText, out decoded, out output);
            }
            catch (BinariexException exc)
            {
                throw exc.AddReaderPosition($@"[{Path.GetFileName(path)}]{cell.FullAddress}");
            }

            base.StepCursor();
        }

        public void Dispose()
        {
            this.package.Dispose();
        }

        void GetDecodedValue(LeafInfo leafInfo, string cellText, out object decoded, out object output)
        {
            if (cellText.EndsWith(".."))
            {
                if (leafInfo.Type == "bin")
                {
                    decoded = Enumerable.Repeat(Convert.ToByte(cellText.Substring(0, 2), 16), leafInfo.Size).ToArray();
                    output = null;
                    return;
                }
                else
                {
                    throw new InvalidDataException();
                }
            }

            var m = Regex.Match(cellText, @"^(?:(.*) )?<([0-9A-Fa-f]+)>$");
            if (m.Success)
            {
                decoded = m.Groups[1].Value ?? "";
                output = GetByteArrayFromBinString(m.Groups[2].Value);
                return;
            }
            output = null;

            switch (leafInfo.Type)
            {
                case "bin":
                    decoded = GetByteArrayFromBinString(cellText);
                    break;
                case "char":
                    decoded = cellText;
                    break;
                case "int":
                case "pbcd":
                    try
                    {
                        decoded = Int64.Parse(cellText);
                    }
                    catch (Exception exc)
                    {
                        throw new BinariexException(exc, "reading input file", "Invalid number format: {0}", cellText != null ? cellText : "''");
                    }
                    break;
                case "uint":
                    try
                    {
                        decoded = UInt64.Parse(cellText);
                    }
                    catch (Exception exc)
                    {
                        throw new BinariexException(exc, "reading input file", "Invalid number format: {0}", cellText != null ? cellText : "''");
                    }
                    break;
                default:
                    throw new InvalidDataException();
            }

        }

        byte[] GetByteArrayFromBinString(string binStr)
        {
            if (binStr.Length % 2 == 1)
            {
                throw new InvalidDataException();
            }

            return Enumerable.Range(0, binStr.Length / 2)
                .Select(i => Convert.ToByte(binStr.Substring(i * 2, 2), 16))
                .ToArray();
        }
    }
}
