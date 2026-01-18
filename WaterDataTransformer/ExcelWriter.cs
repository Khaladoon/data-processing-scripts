using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace WaterBalanceTransformer.Core
{
    public static class ExcelWriter
    {
        public static void WriteAggregated(
            Excel.Worksheet sheet,
            List<AggregatedRecord> data)
        {
            int row = 3;
            foreach (var item in data)
            {
                sheet.Cells[row, 1] = item.Period;
                sheet.Cells[row, 2] = item.Value;
                row++;
            }
        }
    }
}
