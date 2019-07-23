using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace binariex
{
    class ExcelReader : ExcelBase, IReader, IDisposable
    {
        private ExcelPackage package;

        public ExcelReader(string path)
        {
            this.package = new ExcelPackage(new FileInfo(path));
        }

        protected override ExcelWorksheet OpenSheet(string name)
        {
            return package.Workbook.Worksheets[name];
        }

        public void GetValue(LeafInfo leafInfo, out object raw, out object decoded)
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

            var cellText = ctx.Sheet.Cells[ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex, ctx.CursorColumnIndex].Text;
            raw = cellText;

            var valueStr = cellText.EndsWith(">") ? cellText.Substring(0, cellText.LastIndexOf("<") - 1) : cellText;
            switch (leafInfo.Type)
            {
                case "bin":
                    if (valueStr.EndsWith(".."))
                    {
                        decoded = Enumerable.Repeat((byte)0, leafInfo.Size).ToArray();
                    }
                    else
                    {
                        decoded = Enumerable.Range(0, valueStr.Length / 2)
                            .Select(i => Convert.ToByte(valueStr.Substring(i * 2, 2), 16))
                            .ToArray();
                    }
                    break;
                case "char":
                    decoded = valueStr;
                    break;
                case "int":
                case "pbcd":
                    decoded = Int64.Parse(valueStr);
                    break;
                case "uint":
                    decoded = UInt64.Parse(valueStr);
                    break;
                default:
                    throw new InvalidDataException();
            }

            base.StepCursor();
        }

        public void Dispose()
        {
            this.package.Dispose();
        }
    }
}
