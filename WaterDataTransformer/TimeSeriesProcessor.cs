using System;
using System.Collections.Generic;
using System.Linq;

namespace WaterBalanceTransformer.Core
{
    public static class TimeSeriesProcessor
    {
        public static List<AggregatedRecord> AggregateMonthly(List<DailyRecord> data)
        {
            return data
                .GroupBy(d => new { d.Date.Year, d.Date.Month })
                .Select(g => new AggregatedRecord
                {
                    Period = $"{g.Key.Month}-{g.Key.Year}",
                    Value = g.Sum(x => x.Value)
                })
                .ToList();
        }

        public static List<AggregatedRecord> AggregateYearly(List<DailyRecord> data)
        {
            return data
                .GroupBy(d => d.Date.Year)
                .Select(g => new AggregatedRecord
                {
                    Period = g.Key.ToString(),
                    Value = g.Sum(x => x.Value)
                })
                .ToList();
        }

        public static List<AggregatedRecord> AggregateHydrologicalYear(
            List<DailyRecord> data, int startMonth = 10)
        {
            return data
                .GroupBy(d =>
                    d.Date.Month >= startMonth
                        ? $"{d.Date.Year}-{d.Date.Year + 1}"
                        : $"{d.Date.Year - 1}-{d.Date.Year}"
                )
                .Select(g => new AggregatedRecord
                {
                    Period = g.Key,
                    Value = g.Sum(x => x.Value)
                })
                .ToList();
        }
    }
}
