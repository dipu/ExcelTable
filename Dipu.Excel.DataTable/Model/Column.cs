namespace Dipu.Excel.DataTable
{
    using System;

    public class Column<T>
    {
        public Func<T, object> ToExcel;

        public Action<T, object> ToDomain;
    }
}
