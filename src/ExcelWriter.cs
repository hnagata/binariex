using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace binariex
{
    class ExcelWriter : IWriter, IDisposable
    {
        public const string HEADER_MARKER = "HEADER";
        const int MAX_NUM_CHAR_IN_LINE = 254;

        readonly dynamic settings;

        readonly FileInfo outputFile;
        readonly ExcelPackage package;

        readonly Stack<GroupContext> groupCtxStack = new Stack<GroupContext>();
        readonly Dictionary<string, SheetContext> sheetCtxMap = new Dictionary<string, SheetContext>();

        public ExcelWriter(string path, dynamic settings)
        {
            this.settings = settings;
            this.outputFile = new FileInfo(path);
            this.package = new ExcelPackage();
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

                if (ctx.CursorRowIndex > ctx.TopRowIndex && ctx.ColumnCount > 0)
                {
                    var top = ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex;
                    var left = ctx.CursorColumnIndex;
                    var bottom = ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex - 1;
                    var right = ctx.LeftColumnIndex + ctx.ColumnCount - 1;
                    ctx.Sheet.Cells[top, left, bottom, left].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ctx.Sheet.Cells[top, right, bottom, right].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

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
                if (ctx.RowCount > 0 && ctx.ColumnCount > 0)
                {
                    ctx.Sheet.Cells[
                        ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex + ctx.RowCount - 1,
                        ctx.LeftColumnIndex,
                        ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex + ctx.RowCount - 1,
                        ctx.LeftColumnIndex + ctx.ColumnCount - 1
                    ].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

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

            if (ctx.RowCount > 0 && ctx.ColumnCount > 0)
            {
                var top = ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex;
                var left = ctx.LeftColumnIndex;
                var bottom = ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex + ctx.RowCount - 1;
                var right = ctx.LeftColumnIndex + ctx.ColumnCount - 1;
                ctx.Sheet.Cells[bottom, left, bottom, right].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ctx.Sheet.Cells[top, left, bottom, left].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ctx.Sheet.Cells[top, right, bottom, right].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
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
                TopRowIndex = (sheet.Dimension != null ? sheet.Dimension.Rows : 0) - sheetCtx.HeaderRowCount + 1,
                LeftColumnIndex = 1,
                CursorRowIndex = (sheet.Dimension != null ? sheet.Dimension.Rows : 0) - sheetCtx.HeaderRowCount + 1,
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
            var dispRepr = totalRepr.Length > MAX_NUM_CHAR_IN_LINE ? totalRepr.Substring(0, MAX_NUM_CHAR_IN_LINE - 2) + ".." : totalRepr;

            var cell = ctx.Sheet.Cells[ctx.CursorRowIndex + ctx.SheetContext.HeaderRowCount, ctx.CursorColumnIndex];
            cell.Value = dispRepr;

            if (ctx.RepeatEnabled)
            {
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
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

        public void Save()
        {
            foreach (var sheetCtx in this.sheetCtxMap.Values)
            {
                var sheet = sheetCtx.Sheet;

                sheet.Cells.Style.Font.Name = this.settings["fontFamily"];

                if (this.settings["enableAutoFilter"] == "true")
                {
                    sheet.Cells[sheetCtx.HeaderRowCount, 1, sheet.Dimension.Rows, sheet.Dimension.Columns].AutoFilter = true;
                }

                var headerRange = sheet.Cells[1, 1, sheetCtx.HeaderRowCount, sheet.Dimension.Columns];
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.FromName(this.settings["headerColor"]));

                sheet.Cells[1, 1, 1, sheet.Dimension.Columns].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, sheetCtx.HeaderRowCount, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, sheet.Dimension.Columns, sheetCtx.HeaderRowCount, sheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                for (var c = 1; c <= sheet.Dimension.Columns; c++)
                {
                    var appearInCol = false;
                    for (var r = 1; r <= sheetCtx.HeaderRowCount; r++)
                    {
                        appearInCol |= sheet.Cells[r, c].Text != "";
                        if (sheet.Cells[r + 1, c].Text != "")
                        {
                            sheet.Cells[r, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                        else if (c > 1 && !appearInCol)
                        {
                            sheet.Cells[r, c].Style.Border.Bottom.Style = sheet.Cells[r, c - 1].Style.Border.Bottom.Style;
                        }
                        if (sheet.Cells[r, c + 1].Text != "")
                        {
                            sheet.Cells[r, c].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }
                        else if (r > 1)
                        {
                            sheet.Cells[r, c].Style.Border.Right.Style = sheet.Cells[r - 1, c].Style.Border.Right.Style;
                        }
                    }
                }

                var rightOutColumnIndex = sheet.Dimension.Columns + 1;
                sheet.Cells[1, rightOutColumnIndex, sheetCtx.HeaderRowCount, rightOutColumnIndex].Value = HEADER_MARKER;
            }
            this.package.SaveAs(outputFile);
        }

        public void Seek(long offset)
        {
            // Nothing to do
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
