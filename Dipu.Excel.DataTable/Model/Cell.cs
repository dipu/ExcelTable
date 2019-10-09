using Microsoft.Office.Interop.Excel;
using System;

namespace Dipu.Excel.DataTable
{
    public class Cell<T>
    {
        // ReSharper disable once InconsistentNaming
        private object value;

        public IFormatter Formatter { get; set; }

        internal IFormatter PreviousFormatter { get; set; }

        internal Cell(Row<T> row, Column<T> column)
        {
            this.Row = row;
            this.Column = column;
        }

        public Row<T> Row { get; }

        public Column<T> Column { get; }

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
        public void Read()
        {
            var model = this.Row.Model;
            this.Value = this.Column.Read(model);
            if (this.Column.Format != null)
            {
                this.Formatter = this.Column.Format(this);
            }
        }

        public DipuError Write(object value)
        {
            if (this.Column.Write != null)
            {
                if (this.Column.Write(this.Row.Model, value))
                {
                    this.Value = value;
                    return new DipuError() { };
                }
                else
                {
                    // No CanWrite -> Security
                    this.Row.Table.Reset(this);
                    return new DipuError() { NotAutorised = true };
                }

            }
            else
            {
                // No Write defined, so this is a Derived Property
                this.Row.Table.Reset(this);
                return new DipuError() { IsReadOnly = true };
            }
        }

        public int ColumnIndex => this.Row.Table.StartColumn + Array.IndexOf(this.Row.Table.Columns, this.Column);
    }
}
