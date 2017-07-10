using Exchange.WebServices.Extensions.Extensions;
using Microsoft.Exchange.WebServices.Data;
using System;
using Xunit;

namespace Exchange.WebServices.Extensions.Tests
{
    public class DateTimeRelativeExtensionsTest
    {
        [Theory]
        [InlineData(DayOfTheWeek.Monday, DayOfTheWeekIndex.First, "2017-01-02", "2017-02-06")]
        [InlineData(DayOfTheWeek.Saturday, DayOfTheWeekIndex.Second, "2017-02-24", "2017-03-11")]
        [InlineData(DayOfTheWeek.Thursday, DayOfTheWeekIndex.Third, "2017-02-24", "2017-03-16")]
        [InlineData(DayOfTheWeek.Tuesday, DayOfTheWeekIndex.Fourth, "2017-02-24", "2017-03-28")]
        [InlineData(DayOfTheWeek.WeekendDay, DayOfTheWeekIndex.First, "2017-02-24", "2017-03-04")]
        [InlineData(DayOfTheWeek.Friday, DayOfTheWeekIndex.Last, "2017-02-24", "2017-03-31")]
        [InlineData(DayOfTheWeek.Day, DayOfTheWeekIndex.Last, "2017-02-24", "2017-03-31")]
        [InlineData(DayOfTheWeek.Weekday, DayOfTheWeekIndex.Last, "2017-02-24", "2017-03-31")]
        public void ShouldReturnNextMonth(DayOfTheWeek day, DayOfTheWeekIndex index, string from, string expected)
        {
            var date = DateTime.Parse(from);
            var expectedDate = DateTime.Parse(expected);
            var actual = date.AddRelativeMonths(1, day, index);
            Assert.Equal(expectedDate, actual);
        }

        [Theory]
        [InlineData(Month.April, DayOfTheWeek.Monday, DayOfTheWeekIndex.First, "2018-04-02")]
        [InlineData(Month.April, DayOfTheWeek.Saturday, DayOfTheWeekIndex.Second, "2018-04-14")]
        [InlineData(Month.April, DayOfTheWeek.Thursday, DayOfTheWeekIndex.Third, "2018-04-19")]
        [InlineData(Month.April, DayOfTheWeek.Tuesday, DayOfTheWeekIndex.Fourth, "2018-04-24")]
        [InlineData(Month.April, DayOfTheWeek.WeekendDay, DayOfTheWeekIndex.First, "2018-04-01")]
        [InlineData(Month.April, DayOfTheWeek.Friday, DayOfTheWeekIndex.Last, "2018-04-27")]
        [InlineData(Month.April, DayOfTheWeek.Day, DayOfTheWeekIndex.Last, "2018-04-30")]
        [InlineData(Month.April, DayOfTheWeek.Weekday, DayOfTheWeekIndex.Last, "2018-04-30")]
        public void ShouldReturnNextYear(Month month, DayOfTheWeek day, DayOfTheWeekIndex index, string expected)
        {
            var date = DateTime.Parse("2017-01-01");
            var expectedDate = DateTime.Parse(expected);
            var actual = date.AddRelativeYears(1, month, day, index);
            Assert.Equal(expectedDate, actual);
        }
    }
}
