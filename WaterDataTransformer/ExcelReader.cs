using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace WaterBalanceTransformer.Core
{
    public static class ExcelReader
    {
        public static List<DailyRecord> ReadDailyData(
            Excel.Worksheet sheet,
            DateTime startDate)
        {
            var list = new List<DailyRecord>();
            int row = 1;

            while (((Excel.Range)sheet.Cells[row, 1]).Value2 != null)
            {
                double dayIndex = ((Excel.Range)sheet.Cells[row, 1]).Value2;
                double value = ((Excel.Range)sheet.Cells[row, 2]).Value2;

                list.Add(new DailyRecord
                {
                    Date = startDate.AddDays(dayIndex - 1),
                    Value = value
                });

                row++;
            }

            return list;
        }
    }
}
