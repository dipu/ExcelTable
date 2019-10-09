namespace Dipu.Excel.DataTable
{
    using System;

    public class Column<T>
    {
        public Func<T, object> Read;

        public Func<T, object, bool> Write;

        public Func<Cell<T>, IFormatter> Format;
    }
}
