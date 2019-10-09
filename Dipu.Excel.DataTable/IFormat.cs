using Microsoft.Office.Interop.Excel;

namespace Dipu.Excel.DataTable
{
    public interface IFormatter
    {
        void Format(Range range);
    }
}
