using Microsoft.Exchange.WebServices.Data;
using System;

namespace Exchange.WebServices.OccurrenceData.Extensions
{
    public static class DateTimeRelativeExtensions
    {
        public static DateTime AddRelativeMonths(this DateTime datetime, int value, DayOfTheWeek dayOfTheWeek, DayOfTheWeekIndex dayOfTheWeekIndex)
        {
            return Calculate(dayOfTheWeek, dayOfTheWeekIndex, datetime.Month + value, datetime.Year);
        }

        public static DateTime AddRelativeYears(this DateTime datetime, int value, Month month, DayOfTheWeek dayOfTheWeek, DayOfTheWeekIndex dayOfTheWeekIndex)
        {
            return Calculate(dayOfTheWeek, dayOfTheWeekIndex, (int)month, datetime.Year + value);
        }

        private static DateTime Calculate(DayOfTheWeek dayOfTheWeek, DayOfTheWeekIndex dayOfTheWeekIndex, int month, int year)
        {
            if (dayOfTheWeekIndex == DayOfTheWeekIndex.Last)
            {
                int daysInMonth = DateTime.DaysInMonth(year, month);
                var lastDayOfMonth = new DateTime(year, month, daysInMonth);

                while (!CheckDayOfWeek(lastDayOfMonth.DayOfWeek, dayOfTheWeek))
                {
                    lastDayOfMonth = lastDayOfMonth.AddDays(-1);
                }

                return lastDayOfMonth;
            }
            else
            {
                int count = (int)dayOfTheWeekIndex;

                var day = new DateTime(year, month, 1);

                do
                {
                    if (CheckDayOfWeek(day.DayOfWeek, dayOfTheWeek))
                    {
                        if (count != 0)
                        {
                            count--;
                        }
                        else
                        {
                            return day;
                        }
                    }
                    day = day.AddDays(1);

                } while (true);
            }
        }

        private static bool CheckDayOfWeek(DayOfWeek value, DayOfTheWeek expected)
        {
            if (expected == DayOfTheWeek.Day)
            {
                return true;
            }
            if ((DayOfTheWeek)value == expected)
            {
                return true;
            }
            if (expected == DayOfTheWeek.Weekday && (value != DayOfWeek.Saturday && value != DayOfWeek.Sunday))
            {
                return true;
            }
            if (expected == DayOfTheWeek.WeekendDay && (value == DayOfWeek.Saturday || value == DayOfWeek.Sunday))
            {
                return true;
            }
            return false;
        }
    }
}
