namespace Dipu.Excel.DataTable
{
    public class Cell<T>
    {
        private object value;

        internal Cell(Row<T> row)
        {
            this.Row = row;
        }

        public Row<T> Row { get; }

        public object Value
        {
            get => this.value;
            set
            {
                if (!Equals(this.value, value))
                {
                    this.value = value;
                    this.Row.IsDirty = true;
                }
            }
        }

        public void Bind<T>(T model, Column<T> column)
        {
            this.Value = column.ToExcel(model);
        }
    }
}
