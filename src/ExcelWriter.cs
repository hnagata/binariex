using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace binariex
{
    class ExcelWriter : IWriter, IDisposable
    {
        public const string HEADER_MARKER = "HEADER";

        readonly FileInfo outputFile;
        readonly ExcelPackage package;

        readonly Stack<GroupContext> groupCtxStack = new Stack<GroupContext>();
        readonly Dictionary<string, SheetContext> sheetCtxMap = new Dictionary<string, SheetContext>();

        public ExcelWriter(string path)
        {
            this.outputFile = new FileInfo(path);
            this.package = new ExcelPackage();
        }

        public void BeginRepeat()
        {
            var ctx = this.groupCtxStack.Peek();
            ctx.RepeatEnabled = true;
        }

        public void EndRepeat()
        {
            var ctx = this.groupCtxStack.Peek();
            ctx.CursorRowIndex = ctx.TopRowIndex;
            ctx.CursorColumnIndex = ctx.LeftColumnIndex + ctx.ColumnCount;
        }

        public void PopGroup()
        {
            var ctx = this.groupCtxStack.Pop();
            var parentCtx = this.groupCtxStack.Peek();
            if (parentCtx.RepeatEnabled)
            {
                SetBorder(parentCtx, ctx.TopRowIndex, ctx.RowCount, ctx.LeftColumnIndex, ctx.ColumnCount);

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
            this.groupCtxStack.Pop();
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
                SheetContext = parentCtx.SheetContext,
                HeaderRowIndex = parentCtx.HeaderRowIndex + 1,
                TopRowIndex = parentCtx.CursorRowIndex,
                LeftColumnIndex = parentCtx.CursorColumnIndex,
                CursorRowIndex = parentCtx.CursorRowIndex,
                CursorColumnIndex = parentCtx.CursorColumnIndex
            };

            SetHeaderName(newCtx, newCtx.HeaderRowIndex, newCtx.LeftColumnIndex, name);

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

            var sheet = this.package.Workbook.Worksheets[name] ?? this.package.Workbook.Worksheets.Add(name);
            if (!this.sheetCtxMap.TryGetValue(name, out var sheetCtx))
            {
                sheetCtx = new SheetContext
                {
                    Sheet = sheet
                };
                this.sheetCtxMap[name] = sheetCtx;
            }

            var groupCtx = new GroupContext
            {
                Sheet = sheet,
                SheetContext = sheetCtx,
                HeaderRowIndex = 0,
                TopRowIndex = (sheet.Dimension != null ? sheet.Dimension.End.Row : 0) - sheetCtx.HeaderRowCount + 1,
                LeftColumnIndex = 1,
                CursorRowIndex = (sheet.Dimension != null ? sheet.Dimension.End.Row : 0) - sheetCtx.HeaderRowCount + 1,
                CursorColumnIndex = 1,
            };

            this.groupCtxStack.Push(groupCtx);
        }

        public void SetValue(LeafInfo leafInfo, object raw, object decoded)
        {
            if (this.groupCtxStack.Count == 0)
            {
                throw new InvalidOperationException();
            }
            var ctx = this.groupCtxStack.Peek();

            SetHeaderName(ctx, ctx.HeaderRowIndex + 1, ctx.CursorColumnIndex, leafInfo.Name);

            var rawRepr = string.Join("", (raw as byte[]).Select(e => e.ToString("X2")));
            var totalRepr = raw != decoded ? $@"{decoded.ToString()} <{rawRepr}>" : rawRepr;

            var cell = ctx.Sheet.Cells[ctx.CursorRowIndex + ctx.SheetContext.HeaderRowCount, ctx.CursorColumnIndex];
            cell.Value = totalRepr;

            if (ctx.RepeatEnabled)
            {
                SetBorder(ctx, ctx.CursorRowIndex, ctx.CursorColumnIndex, 1, 1);
            }

            if (ctx.RepeatEnabled)
            {
                ctx.CursorRowIndex += 1;
                ctx.RowCount = Math.Max(ctx.RowCount, ctx.CursorRowIndex - ctx.TopRowIndex);
                ctx.ColumnCount = Math.Max(ctx.ColumnCount, 1);
            }
            else
            {
                ctx.CursorColumnIndex += 1;
                ctx.RowCount = Math.Max(ctx.RowCount, 1);
                ctx.ColumnCount += 1;
            }
        }

        public void Save()
        {
            foreach (var sheetCtx in this.sheetCtxMap.Values)
            {
                int rightOutColumnIndex = sheetCtx.Sheet.Dimension.End.Column + 1;
                sheetCtx.Sheet.Cells[1, rightOutColumnIndex, sheetCtx.HeaderRowCount, rightOutColumnIndex].Value = HEADER_MARKER;
            }
            this.package.SaveAs(outputFile);
        }

        public void Dispose()
        {
            this.package.Dispose();
        }

        void SetHeaderName(GroupContext ctx, int headerRowIndex, int columnIndex, string name)
        {
            if (headerRowIndex > ctx.SheetContext.HeaderRowCount)
            {
                ctx.Sheet.InsertRow(headerRowIndex, 1);
                ctx.SheetContext.HeaderRowCount += 1;
            }

            var headerCell = ctx.Sheet.Cells[headerRowIndex, columnIndex];
            if (headerCell.Value == null)
            {
                headerCell.Value = name;
            }
            else if (headerCell.Text != name)
            {
                throw new InvalidOperationException();
            }
        }

        void SetBorder(GroupContext ctx, int topRowIndex, int rowCount, int leftColumnIndex, int columnCount)
        {
            if (rowCount == 0 || columnCount == 0)
            {
                return;
            }
            var range = ctx.Sheet.Cells[
                topRowIndex + ctx.SheetContext.HeaderRowCount, 
                leftColumnIndex, 
                topRowIndex + ctx.SheetContext.HeaderRowCount + rowCount - 1, 
                leftColumnIndex + columnCount - 1];
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        class GroupContext
        {
            public ExcelWorksheet Sheet { get; set; }
            public SheetContext SheetContext { get; set; }
            public int HeaderRowIndex { get; set; }
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
            public int HeaderRowCount { get; set; }
        }
    }
}
