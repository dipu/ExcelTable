namespace Dipu.Excel.DataTable
{
    using System.Collections.Generic;

    public class Row<T>
    {
        public Row()
        {
            this.Cells = new List<Cell<T>>();
        }

        public List<Cell<T>> Cells { get; set; }

        public bool IsDirty { get; set; }

        public void Bind(T model, Column<T>[] columns)
        {
            for (var i = 0; i < columns.Length; i++)
            {
                if (this.Cells.Count == i)
                {
                    this.Cells.Add(new Cell<T>(this));
                }

                var cell = this.Cells[i];
                cell.Bind(model, columns[i]);
            }

            // Remove superfluous cells
        }
    }
}
