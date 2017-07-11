using Exchange.WebServices.Extensions.Extensions;
using Microsoft.Exchange.WebServices.Data;
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
            Recurrence reccurrence,
            Occurrence defaultOccurrence,
            Expression<Func<DateTime, DateTime>> predicate)
        {
            var result = new OccurrenceCollection();

            var index = 1;

            var start = defaultOccurrence.Start;
            var end = defaultOccurrence.End;

            if (reccurrence.EndDate != null)
            {
                while (start <= reccurrence.EndDate)
                {
                    result.Add(Generate(defaultOccurrence, predicate, ref index, ref start, ref end));
                }
            }
            else
            {
                for (int i = 0; i < (reccurrence.NumberOfOccurrences ?? 0); i++)
                {
                    result.Add(Generate(defaultOccurrence, predicate, ref index, ref start, ref end));
                }
            }

            return result;
        }

        private static Occurrence Generate(Occurrence defaultOccurrence, Expression<Func<DateTime, DateTime>> predicate, ref int index, ref DateTime start, ref DateTime end)
        {
            var occurrence = new Occurrence(defaultOccurrence)
            {
                Start = start,
                End = end,
                Index = index
            };

            start = predicate.Compile()(start);
            end = predicate.Compile()(end);
            index++;

            return occurrence;
        }

        public static OccurrenceCollection Convert(DailyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddDays(pattern.Interval);
            return Convert(pattern, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(WeeklyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddDays(pattern.Interval * 7);
            return Convert(pattern, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(MonthlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddMonths(pattern.Interval);
            return Convert(pattern, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(YearlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate.AddYears(1);
            return Convert(pattern, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(RelativeMonthlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate
                .AddRelativeMonths(pattern.Interval, pattern.DayOfTheWeek, pattern.DayOfTheWeekIndex);

            return Convert(pattern, defaultOccurrence, predicate);
        }

        public static OccurrenceCollection Convert(RelativeYearlyPattern pattern, Occurrence defaultOccurrence)
        {
            Expression<Func<DateTime, DateTime>> predicate = previousDate => previousDate
                .AddRelativeYears(1, pattern.Month, pattern.DayOfTheWeek, pattern.DayOfTheWeekIndex);

            return Convert(pattern, defaultOccurrence, predicate);
        }
    }
}
