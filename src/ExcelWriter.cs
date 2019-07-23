using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace binariex
{
    class ExcelWriter : ExcelBase, IWriter, IDisposable
    {
        const int MAX_NUM_CHAR_IN_LINE = 254;

        readonly dynamic settings;
        readonly FileInfo outputFile;

        readonly ExcelPackage package;

        public ExcelWriter(string path, dynamic settings)
        {
            this.outputFile = new FileInfo(path);
            this.settings = settings;

            this.package = new ExcelPackage();
        }

        public override void EndRepeat()
        {
            var ctx = this.CurrentContext;
            if (ctx != null && ctx.CursorRowIndex > ctx.TopRowIndex && ctx.ColumnCount > 0)
            {
                var top = ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex;
                var left = ctx.CursorColumnIndex;
                var bottom = ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex - 1;
                var right = ctx.LeftColumnIndex + ctx.ColumnCount - 1;
                ctx.Sheet.Cells[top, left, bottom, left].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ctx.Sheet.Cells[top, right, bottom, right].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            base.EndRepeat();
        }

        public override void PopGroup()
        {
            if (this.ParentContext.RepeatEnabled)
            {
                var ctx = this.CurrentContext;
                if (ctx.RowCount > 0 && ctx.ColumnCount > 0)
                {
                    ctx.Sheet.Cells[
                        ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex + ctx.RowCount - 1,
                        ctx.LeftColumnIndex,
                        ctx.SheetContext.HeaderRowCount + ctx.TopRowIndex + ctx.RowCount - 1,
                        ctx.LeftColumnIndex + ctx.ColumnCount - 1
                    ].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
            }
            base.PopGroup();
        }

        public override void PopSheet()
        {
            var ctx = this.CurrentContext;
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
            base.PopSheet();
        }

        public override void PushGroup(string name)
        {
            base.PushGroup(name);

            var ctx = this.CurrentContext;
            SetHeaderName(ctx, ctx.HeaderRowIndex, ctx.LeftColumnIndex, name);
        }

        protected override ExcelWorksheet OpenSheet(string name)
        {
            return package.Workbook.Worksheets[name] ?? package.Workbook.Worksheets.Add(name);
        }

        public void SetValue(LeafInfo leafInfo, object raw, object decoded)
        {
            var ctx = this.CurrentContext;
            if (ctx == null)
            {
                throw new InvalidOperationException();
            }

            SetHeaderName(ctx, ctx.HeaderRowIndex + 1, ctx.CursorColumnIndex, leafInfo.Name);

            var rawRepr = string.Join("", (raw as byte[]).Select(e => e.ToString("X2")));
            var totalRepr = raw != decoded ? $@"{decoded.ToString()} <{rawRepr}>" : rawRepr;
            var dispRepr = totalRepr.Length > MAX_NUM_CHAR_IN_LINE ? totalRepr.Substring(0, MAX_NUM_CHAR_IN_LINE - 2) + ".." : totalRepr;

            var cell = ctx.Sheet.Cells[ctx.SheetContext.HeaderRowCount + ctx.CursorRowIndex, ctx.CursorColumnIndex];
            cell.Value = dispRepr;

            if (ctx.RepeatEnabled)
            {
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            base.StepCursor();
        }

        public void Save()
        {
            foreach (var sheetCtx in this.SheetContextMap.Values)
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
    }
}
