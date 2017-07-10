using Exchange.WebServices.Extensions.Extensions;
using System;
using System.Linq.Expressions;
using static Microsoft.Exchange.WebServices.Data.Recurrence;

namespace Exchange.WebServices.Extensions
{
    public static class PatternConverter
    {
        // YearlyPattern
        // RelativeMonthlyPattern
        // RelativeYearlyPattern
        public static OccurrenceCollection Convert(
            DateTime endDate,
            Occurrence defaultOccurrence,
            Expression<Func<DateTime, DateTime>> predicate)
        {
            var result = new OccurrenceCollection();

            var index = 1;

            var start = defaultOccurrence.Start;
            var end = defaultOccurrence.End;

            while (start <= endDate)
            {
                var occurrence = new Occurrence(defaultOccurrence)
                {
                    Start = start,
                    End = end,
                    Index = index
                };

                result.Add(occurrence);

                start = predicate.Compile()(start);
                end = predicate.Compile()(end);
                index++;
            }

            return result;
        }

        public static OccurrenceCollection Convert(DailyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddDays(pattern.Interval);
            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(WeeklyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddDays(pattern.Interval * 7);
            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(MonthlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddMonths(pattern.Interval);
            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(YearlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddYears(1);
            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(RelativeMonthlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate
                .AddRelativeMonths(pattern.Interval, pattern.DayOfTheWeek, pattern.DayOfTheWeekIndex);

            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(RelativeYearlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate
                .AddRelativeYears(1, pattern.Month, pattern.DayOfTheWeek, pattern.DayOfTheWeekIndex);

            return Convert(pattern.EndDate.Value, defaultOccurrence, predicate);
        }
    }
}
