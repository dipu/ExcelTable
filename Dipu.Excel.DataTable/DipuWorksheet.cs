namespace Dipu.Excel.DataTable
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;

    public class DipuWorksheet : IDisposable
    {
        public DipuWorksheet(Worksheet worksheet)
        {
            this.Worksheet = worksheet;
        }

        ~DipuWorksheet()
        {
            this.Dispose(false);
        }

        public Worksheet Worksheet { get; private set; }

        public DipuRange CreateRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            return new DipuRange(this.Worksheet, fromRow, fromColumn, toRow, toColumn);
        }
        

        public void Dispose()
        {
            Marshal.FinalReleaseComObject(this.Worksheet);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (this.Worksheet != null)
            {
                Marshal.FinalReleaseComObject(this.Worksheet);
                this.Worksheet = null;
            }
        }
    }
}
