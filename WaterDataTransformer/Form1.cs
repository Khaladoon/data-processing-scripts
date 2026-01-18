using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WaterBalanceTransformer.Core;

namespace WaterBalanceTransformer.UI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            var xl = new Excel.Application();
            var wb = xl.Workbooks.Open(@"input.xlsx");
            var ws = (Excel.Worksheet)wb.Worksheets[1];

            var daily = ExcelReader.ReadDailyData(ws, dateTimePicker1.Value);

            var monthly = TimeSeriesProcessor.AggregateMonthly(daily);
            var yearly = TimeSeriesProcessor.AggregateYearly(daily);
            var hydro = TimeSeriesProcessor.AggregateHydrologicalYear(daily);

            var outWb = xl.Workbooks.Add();
            ExcelWriter.WriteAggregated(
                (Excel.Worksheet)outWb.Worksheets[1], monthly);

            outWb.SaveAs(@"Output.xlsx");
            xl.Quit();
        }
    }
}
