namespace WaterBalanceTransformer.Core
{
    public class DailyRecord
    {
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    public class AggregatedRecord
    {
        public string Period { get; set; }
        public double Value { get; set; }
    }
}
