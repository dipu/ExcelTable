namespace Dipu.Excel.DataTable
{
    using System.Collections.Generic;
    using System.Linq;

    public class Table<T>
    {
        /// <summary>
        /// Creates an Table, representing an Excel Range. startRow and startColumn are 1-based (like excel)
        /// </summary>
        /// <param name="allorsWorksheet"></param>
        /// <param name="columns"></param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        public Table(AllorsWorksheet allorsWorksheet, Column<T>[] columns, int startRow, int startColumn)
        {
            this.AllorsWorksheet = allorsWorksheet;
            this.Columns = columns;
            this.StartRow = startRow;
            this.StartColumn = startColumn;
            this.Rows = new List<Row<T>>();
        }

        public AllorsWorksheet AllorsWorksheet { get; }

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

        public void Bind(IEnumerable<T> data)
        {
            var dataModel = data.ToArray();
            int i = 0;
            foreach (var model in dataModel)
            {
                if (this.Rows.Count == i)
                {
                    this.Rows.Add(new Row<T>());
                }

                var row = this.Rows[i];
                row.Bind(model, this.Columns);

                ++i;
            }

            // Remove superfluous rows
            if (this.Rows.Count > dataModel.Count())
            {

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
                
                using (var allorsRange = this.AllorsWorksheet.CreateRange(this.StartRow + startRow, startColumn, this.StartRow + endRow, endColumn))
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

                    allorsRange.Range.Value2 = array;
                }
            }

            return ranges;
        }
    }
}
