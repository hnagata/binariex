using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace binariex
{
    abstract class ExcelBase
    {
        protected const string HEADER_MARKER = "HEADER";

        readonly Stack<GroupContext> groupCtxStack = new Stack<GroupContext>();
        readonly Dictionary<string, SheetContext> sheetCtxMap = new Dictionary<string, SheetContext>();

        protected GroupContext CurrentContext {
            get
            {
                return groupCtxStack.Count > 0 ? groupCtxStack.Peek() : null;
            }
        }

        protected GroupContext ParentContext
        {
            get
            {
                return groupCtxStack.Count > 1 ? groupCtxStack.Skip(1).First() : null;
            }
        }

        protected IReadOnlyDictionary<string, SheetContext> SheetContextMap
        {
            get
            {
                return this.sheetCtxMap;
            }
        }

        public virtual void BeginRepeat()
        {
            if (this.groupCtxStack.Count > 0)
            {
                var ctx = this.groupCtxStack.Peek();
                ctx.RepeatEnabled = true;
            }
        }

        public virtual void EndRepeat()
        {
            if (this.groupCtxStack.Count > 0)
            {
                var ctx = this.groupCtxStack.Peek();
                ctx.CursorRowIndex = ctx.TopRowIndex;
                ctx.CursorColumnIndex = ctx.LeftColumnIndex + ctx.ColumnCount;
                ctx.RepeatEnabled = false;
            }
        }

        public virtual void PopGroup()
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

        public virtual void PopSheet()
        {
            var ctx = this.groupCtxStack.Pop();
            ctx.SheetContext.CursorRowIndex += ctx.RowCount;
        }

        public virtual void PushGroup(string name)
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

            var sheet = OpenSheet(name);
            if (!this.sheetCtxMap.TryGetValue(name, out var sheetCtx))
            {
                var cursorRowIndex = 1;
                if (sheet.Dimension != null)
                {
                    while (sheet.Cells[cursorRowIndex, sheet.Dimension.Columns].Text == ExcelWriter.HEADER_MARKER)
                    {
                        cursorRowIndex++;
                    }
                }

                sheetCtx = new SheetContext
                {
                    Sheet = sheet,
                    HeaderRowCount = cursorRowIndex - 1,
                    CursorRowIndex = 1
                };
                this.sheetCtxMap[name] = sheetCtx;
            }

            var groupCtx = new GroupContext
            {
                Sheet = sheet,
                SheetContext = sheetCtx,
                HeaderRowIndex = 0,
                TopRowIndex = sheetCtx.CursorRowIndex,
                LeftColumnIndex = 1,
                CursorRowIndex = sheetCtx.CursorRowIndex,
                CursorColumnIndex = 1,
            };

            this.groupCtxStack.Push(groupCtx);
        }

        protected abstract ExcelWorksheet OpenSheet(string name);

        protected void StepCursor()
        {
            var ctx = this.CurrentContext;
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

        public void Seek(long offset)
        {
            // Nothing to do
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

        protected class GroupContext
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

        protected class SheetContext
        {
            public ExcelWorksheet Sheet { get; set; }
            public int HeaderRowCount { get; set; }
            public int CursorRowIndex { get; set; }
        }
    }
}
