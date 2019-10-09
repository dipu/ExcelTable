namespace Dipu.Excel.DataTable
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public class Table<T>
    {
        /// <summary>
        /// Creates an Table, representing an Excel Range. startRow and startColumn are 1-based (like excel)
        /// </summary>
        /// <param name="dipuWorksheet"></param>
        /// <param name="columns"></param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        public Table(DipuWorksheet dipuWorksheet, Column<T>[] columns, int startRow, int startColumn)
        {
            this.DipuWorksheet = dipuWorksheet;
            this.Columns = columns;
            this.StartRow = startRow;
            this.StartColumn = startColumn;
            this.Rows = new List<Row<T>>();
        }

        public DipuWorksheet DipuWorksheet { get; }

        public Column<T>[] Columns { get; }

        /// <summary>
        /// 1-Based StartRow. Points to the Excel Range.Row (Starting Row)
        /// </summary>
        public int StartRow { get; }

        /// <summary>
        /// 1-Based StartColumn. Points to the Excel Range.Column (Starting Column)
        /// </summary>
        public int StartColumn { get; }

        public List<Row<T>> Rows { get; set; }


        public void Read(IEnumerable<T> data)
        {
            var dataModel = data.ToArray();
            int i = 0;

            foreach (var model in dataModel)
            {
                if (this.Rows.Count == i)
                {
                    this.Rows.Add(new Row<T>(this, this.StartRow + i));
                }

                var row = this.Rows[i];
                row.Read(model, this.Columns);

                ++i;
            }

            // Remove superfluous rows
            if (this.Rows.Count > dataModel.Count())
            {

            }
        }

        internal void Reset(Cell<T> cell)
        {
            using (var dipuRange = this.DipuWorksheet.CreateRange(cell.Row.Index, cell.ColumnIndex, cell.Row.Index, cell.ColumnIndex))
            {
                dipuRange.Range.Value2 = cell.Value;
            }
        }

        public IReadOnlyList<int[]> Flush()
        {
            var ranges = new List<int[]>();
            for (var i = 0; i < this.Rows.Count; i++)
            {
                var row = this.Rows[i];
                if (row.IsDirty)
                {
                    if (ranges.Count == 0 || ranges[ranges.Count - 1][1] != 0)
                    {
                        var range = new[] { i, 0 };
                        ranges.Add(range);
                    }
                }
                else
                {
                    if (ranges.Count != 0 && ranges[ranges.Count - 1][1] == 0)
                    {
                        ranges[ranges.Count - 1][1] = i - 1;
                    }
                }

                row.IsDirty = false;
            }

            if (ranges.Count != 0 && ranges[ranges.Count - 1][1] == 0)
            {
                ranges[ranges.Count - 1][1] = this.Rows.Count - 1;
            }

            foreach (var range in ranges)
            {
                // Zero-Based Row
                var startRow = range[0];
                var endRow = range[1];

                // 1-Based Column
                var startColumn = this.StartColumn;
                var endColumn = this.StartColumn + this.Columns.Length - 1;

                using (var dipuRange = this.DipuWorksheet.CreateRange(this.StartRow + startRow, startColumn, this.StartRow + endRow, endColumn))
                {
                    var rowCount = endRow - startRow + 1;
                    var columnCount = this.Columns.Length;

                    var array = new object[rowCount, columnCount];
                    for (var i = 0; i < rowCount; i++)
                    {
                        for (var j = 0; j < columnCount; j++)
                        {
                            array[i, j] = this.Rows[startRow + i].Cells[j].Value;
                        }
                    }

                    dipuRange.Range.Value2 = array;
                }
            }

            this.FlushFormatByColumn();

            return ranges;
        }

        private void FlushFormatByColumn()
        {
            if (this.Rows.Any())
            {
                var fromRowIndex = this.StartRow;
                var fromColumnIndex = this.StartColumn;

                // We treat all rows in the same manner, security on all these cells is the same for all cells in a row
                var toRowIndex = this.StartRow + this.Rows.Count - 1;

                var toColumnIndex = this.StartColumn;

                for (var i = 0; i < this.Columns.Length; i++)
                {
                    var cell = this.Rows[0].Cells[i];

                    Cell<T> previousCell = null;
                    if (i > 0)
                    {
                        previousCell = this.Rows[0].Cells[i - 1];
                    }

                    // Detect a difference in the  column formatting
                    if (previousCell != null && previousCell.Formatter != cell.Formatter)
                    {
                        var formatter = previousCell.Formatter;

                        if (formatter != null)
                        {
                            using (var dipuRange = this.DipuWorksheet.CreateRange(fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex))
                            {
                                formatter.Format(dipuRange.Range);
                            }
                        }

                        cell.PreviousFormatter = cell.Formatter;

                        // Set next starting Point to the next column
                        fromColumnIndex = toColumnIndex + 1;
                        toColumnIndex = fromColumnIndex;
                    }
                    else
                    {
                        // increase the toColumnIndex.
                        toColumnIndex = this.StartColumn + i;
                    }
                }
            }
        }
    }
}
