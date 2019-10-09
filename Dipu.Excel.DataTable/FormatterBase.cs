using Microsoft.Office.Interop.Excel;

namespace Dipu.Excel.DataTable
{
    public abstract class FormatterBase
    {
        public void SetInsideBorders(Range range)
        {
            range.Borders.Item[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlInsideHorizontal].Color = XlRgbColor.rgbBlack;
            range.Borders.Item[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            range.Borders.Item[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlInsideVertical].Color = XlRgbColor.rgbBlack;
            range.Borders.Item[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
        }
    }
}
