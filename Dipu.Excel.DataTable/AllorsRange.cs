namespace Dipu.Excel.DataTable
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;

    public class AllorsRange : IDisposable
    {
        public AllorsRange(Worksheet worksheet, int fromRow, int fromColumn, int toRow, int toColumn)
        {
            Range fromCell = null;
            Range toCell = null;

            try
            {
                fromCell = (Range)worksheet.Cells[fromRow, fromColumn];
                toCell = (Range)worksheet.Cells[toRow, toColumn];
                this.Range = worksheet.Range[fromCell, toCell];
            }
            finally
            {
                if (fromCell != null)
                {
                    Marshal.FinalReleaseComObject(fromCell);
                }

                if (toCell != null)
                {
                    Marshal.FinalReleaseComObject(toCell);
                }
            }
        }

        ~AllorsRange()
        {
            this.Dispose(false);
        }

        public Range Range { get; private set; }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (this.Range != null)
            {
                Marshal.FinalReleaseComObject(this.Range);
                this.Range = null;
            }
        }
    }
}
