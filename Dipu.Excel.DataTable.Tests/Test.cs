namespace Dipu.Excel.DataTable.Tests
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;

    public class Test : IDisposable
    {
        protected Test()
        {
            this.Application = new ApplicationClass { Visible = true, };
            this.Workbook = this.Application.Workbooks.Add();
        }

        public ApplicationClass Application { get; set; }

        public Workbook Workbook { get; set; }

        public void Dispose()
        {
            this.Workbook.Close(false);
            Marshal.FinalReleaseComObject(this.Workbook);
            this.Application.Quit();
            Marshal.FinalReleaseComObject(this.Application);
        }
    }
}
