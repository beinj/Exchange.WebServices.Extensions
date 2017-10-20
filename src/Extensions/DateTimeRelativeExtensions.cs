using Microsoft.Exchange.WebServices.Data;
using System;

namespace Exchange.WebServices.OccurrenceData.Extensions
{
    public static class DateTimeRelativeExtensions
    { 
        /// <summary>
        /// Returns a new <see cref="System.DateTime"/> that adds the relative number of months to the value of this instance.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <param name="value">The value.</param>
        /// <param name="dayOfTheWeek">The day of the week.</param>
        /// <param name="dayOfTheWeekIndex">Index of the day of the week.</param>
        public static DateTime AddRelativeMonths(this DateTime datetime, int value, DayOfTheWeek dayOfTheWeek, DayOfTheWeekIndex dayOfTheWeekIndex)
        {
            var v = datetime.AddMonths(value);
            return Calculate(dayOfTheWeek, dayOfTheWeekIndex, v.Month, v.Year);
        }

        /// <summary>
        /// Returns a new <see cref="System.DateTime"/> that adds the relative number of years to the value of this instance.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <param name="value">The value.</param>
        /// <param name="month">The month.</param>
        /// <param name="dayOfTheWeek">The day of the week.</param>
        /// <param name="dayOfTheWeekIndex">Index of the day of the week.</param>
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
