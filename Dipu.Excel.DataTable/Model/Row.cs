namespace Dipu.Excel.DataTable
{
    using System.Collections.Generic;

    public class Row<T>
    {
        public Row(Table<T> table, int index)
        {
            this.Table = table;
            this.Index = index;
            this.Cells = new List<Cell<T>>();
        }

        public Table<T> Table { get; }

        public int Index { get; }

        public List<Cell<T>> Cells { get; }

        public bool IsDirty { get; set; }

        public T Model { get; private set; }

        public void Read(T model, Column<T>[] columns)
        {
            this.Model = model;

            for (var i = 0; i < columns.Length; i++)
            {
                if (this.Cells.Count == i)
                {
                    this.Cells.Add(new Cell<T>(this, columns[i]));
                }

                var cell = this.Cells[i];
                cell.Bind();
            }

            //TODO Remove superfluous cells
        }
    }
}
