using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace binariex
{
    class ExcelReader : IReader, IDisposable
    {
        readonly ExcelPackage package;

        readonly Stack<GroupContext> groupCtxStack = new Stack<GroupContext>();
        readonly Dictionary<string, SheetContext> sheetCtxMap = new Dictionary<string, SheetContext>();

        public ExcelReader(string path)
        {
            this.package = new ExcelPackage(new FileInfo(path));
        }

        public void BeginRepeat()
        {
            if (this.groupCtxStack.Count > 0)
            {
                var ctx = this.groupCtxStack.Peek();
                ctx.RepeatEnabled = true;
            }
        }

        public void EndRepeat()
        {
            if (this.groupCtxStack.Count > 0)
            {
                var ctx = this.groupCtxStack.Peek();
                ctx.CursorRowIndex = ctx.TopRowIndex;
                ctx.CursorColumnIndex = ctx.LeftColumnIndex + ctx.ColumnCount;
                ctx.RepeatEnabled = false;
            }
        }

        public void PopGroup()
        {
            var ctx = this.groupCtxStack.Pop();
            var parentCtx = this.groupCtxStack.Peek();
            if (parentCtx.RepeatEnabled)
            {
                parentCtx.CursorRowIndex += ctx.RowCount;
                parentCtx.RowCount = Math.Max(parentCtx.RowCount, ctx.TopRowIndex + ctx.RowCount - parentCtx.TopRowIndex);
                parentCtx.ColumnCount = Math.Max(parentCtx.ColumnCount, ctx.LeftColumnIndex + ctx.ColumnCount - parentCtx.LeftColumnIndex);
            }
            else
            {
                parentCtx.CursorColumnIndex += ctx.ColumnCount;
                parentCtx.RowCount = Math.Max(parentCtx.RowCount, ctx.TopRowIndex + ctx.RowCount - parentCtx.TopRowIndex);
                parentCtx.ColumnCount += ctx.ColumnCount;
            }
        }

        public void PopSheet()
        {
            var ctx = this.groupCtxStack.Pop();
            ctx.SheetContext.CursorRowIndex += ctx.RowCount;
        }

        public void PushGroup(string name)
        {
            if (this.groupCtxStack.Count == 0)
            {
                throw new InvalidOperationException();
            }
            var parentCtx = this.groupCtxStack.Peek();

            var newCtx = new GroupContext
            {
                Sheet = parentCtx.Sheet,
                TopRowIndex = parentCtx.CursorRowIndex,
                LeftColumnIndex = parentCtx.CursorColumnIndex,
                CursorRowIndex = parentCtx.CursorRowIndex,
                CursorColumnIndex = parentCtx.CursorColumnIndex
            };

            this.groupCtxStack.Push(newCtx);
        }

        public void PushSheet(string name)
        {
            foreach (var parentCtx in this.groupCtxStack)
            {
                if (parentCtx.Sheet.Name == name)
                {
                    throw new InvalidOperationException();
                }
            }

            var sheet = this.package.Workbook.Worksheets[name];
            if (!this.sheetCtxMap.TryGetValue(name, out var sheetCtx))
            {
                var cursorRowIndex = 1;
                while (sheet.Cells[cursorRowIndex, sheet.Dimension.Columns].Text == ExcelWriter.HEADER_MARKER)
                {
                    cursorRowIndex++;
                }
                sheetCtx = new SheetContext
                {
                    Sheet = sheet,
                    CursorRowIndex = cursorRowIndex
                };
                this.sheetCtxMap[name] = sheetCtx;
            }

            var groupCtx = new GroupContext
            {
                Sheet = sheet,
                SheetContext = sheetCtx,
                TopRowIndex = sheetCtx.CursorRowIndex,
                LeftColumnIndex = 1,
                CursorRowIndex = sheetCtx.CursorRowIndex,
                CursorColumnIndex = 1,
            };

            this.groupCtxStack.Push(groupCtx);
        }

        public void GetValue(LeafInfo leafInfo, out object raw, out object decoded)
        {
            if (this.groupCtxStack.Count == 0)
            {
                throw new InvalidOperationException();
            }
            var ctx = this.groupCtxStack.Peek();

            if (ctx.CursorRowIndex > ctx.Sheet.Dimension.Rows)
            {
                throw new EndOfStreamException();
            }

            var cellText = ctx.Sheet.Cells[ctx.CursorRowIndex, ctx.CursorColumnIndex].Text;
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

            if (ctx.RepeatEnabled)
            {
                ctx.CursorRowIndex += 1;
                ctx.RowCount = Math.Max(ctx.RowCount, ctx.CursorRowIndex - ctx.TopRowIndex);
                ctx.ColumnCount = Math.Max(ctx.ColumnCount, ctx.CursorColumnIndex - ctx.LeftColumnIndex + 1);
            }
            else
            {
                ctx.CursorColumnIndex += 1;
                ctx.RowCount = Math.Max(ctx.RowCount, 1);
                ctx.ColumnCount += 1;
            }
        }

        public void Dispose()
        {
            this.package.Dispose();
        }

        class GroupContext
        {
            public ExcelWorksheet Sheet { get; set; }
            public SheetContext SheetContext { get; set; }
            public int TopRowIndex { get; set; }
            public int LeftColumnIndex { get; set; }
            public int CursorRowIndex { get; set; }
            public int CursorColumnIndex { get; set; }
            public int RowCount { get; set; }
            public int ColumnCount { get; set; }
            public bool RepeatEnabled { get; set; }
        }

        class SheetContext
        {
            public ExcelWorksheet Sheet { get; set; }
            public int CursorRowIndex { get; set; }
        }
    }
}
