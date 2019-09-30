namespace Dipu.Excel.DataTable
{
    public class Cell<T>
    {
        // ReSharper disable once InconsistentNaming
        private object value;

        internal Cell(Row<T> row)
        {
            this.Row = row;
        }

        public Row<T> Row { get; }

        /// <summary>
        /// Get or sets the value.
        /// </summary>
        public object Value
        {
            get => this.value;
            private set
            {
                if (!Equals(this.value, value))
                {
                    this.value = value;
                    this.Row.IsDirty = true;
                }
            }
        }

        /// <summary>
        /// Binds the Value to the result of the Func&lt;T, model&gt; defined in the Column
        /// </summary>
        /// <param name="model"></param>
        /// <param name="column"></param>
        public void Bind(T model, Column<T> column)
        {
            this.Value = column.ToExcel(model);
        }
    }
}
