using Microsoft.Exchange.WebServices.Data;
using System;
using System.Linq;
using Xunit;
using static Microsoft.Exchange.WebServices.Data.Recurrence;

namespace Exchange.WebServices.OccurrenceData.Tests
{
    public class PatternConverterTest
    {
        [Fact]
        public void ShouldContainDailyOccurences()
        {
            var date = DateTime.Now.Date;
            var pattern = new DailyPattern(date, 1)
            {
                EndDate = DateTime.Now.Date.AddDays(5)
            };

            var occurrence = new Occurrence
            {
                Start = date.AddHours(1),
                End = date.AddHours(2)
            };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(6, result.Count);
            Assert.Equal(date.AddHours(1), result.First().Start);
        }

        [Fact]
        public void ShouldContainWeeklyOccurences()
        {
            var date = DateTime.Parse("2017-01-02T00:00:00");

            var pattern = new WeeklyPattern(date, 1)
            {
                EndDate = date.AddDays(14)
            };

            pattern.DaysOfTheWeek.Add(DayOfTheWeek.Monday);
            pattern.DaysOfTheWeek.Add(DayOfTheWeek.Tuesday);
            pattern.DaysOfTheWeek.Add(DayOfTheWeek.Wednesday);

            var occurrence = new Occurrence
            {
                Start = date,
                End = date.AddHours(1)
            };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(7, result.Count);
            Assert.Equal(date, result.First().Start);
        }

        [Fact]
        public void ShouldContainMonthlyOccurences()
        {
            DateTime date = DateTime.Now.Date;
            var pattern = new MonthlyPattern(date, 1, DateTime.Now.Day)
            {
                EndDate = date.AddDays(45)
            };

            var occurrence = new Occurrence { Start = date, End = date.AddHours(1) };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(2, result.Count);
            Assert.Equal(date, result.First().Start);
        }

        [Fact]
        public void ShouldContainYearlyOccurences()
        {
            DateTime date = DateTime.Now.Date;
            var pattern = new YearlyPattern(date, (Month)date.Month, date.Day)
            {
                EndDate = date.AddYears(2)
            };

            var occurrence = new Occurrence { Start = date, End = date.AddHours(1) };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(3, result.Count);
            Assert.Equal(date, result.First().Start);
        }


        [Fact]
        public void ShouldContainRelativeMonthlyOccurences()
        {
            var date = DateTime.Parse("2017-01-01T00:00:00");

            var pattern = new RelativeMonthlyPattern(date, 1, DayOfTheWeek.Sunday, DayOfTheWeekIndex.First)
            {
                EndDate = date.AddMonths(2)
            };

            var occurrence = new Occurrence { Start = date, End = date.AddHours(1) };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(2, result.Count);
            Assert.Equal(date, result.First().Start);
        }

        [Fact]
        public void ShouldContainRelativeYearlyOccurences()
        {
            var date = DateTime.Parse("2017-01-01T00:00:00");

            var pattern = new RelativeYearlyPattern(date, Month.January, DayOfTheWeek.Sunday, DayOfTheWeekIndex.First)
            {
                EndDate = DateTime.Now.Date.AddYears(2)
            };

            var occurrence = new Occurrence { Start = date, End = date.AddHours(1) };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(3, result.Count);
            Assert.Equal(date, result.First().Start);
        }

        [Fact]
        public void ShouldWorkWithNumberOfOccurrences()
        {
            var date = DateTime.Now.Date;
            var pattern = new DailyPattern(date, 1)
            {
                NumberOfOccurrences = 5
            };

            var occurrence = new Occurrence
            {
                Start = date,
                End = date.AddHours(1)
            };

            var result = PatternConverter.Convert(pattern, occurrence);

            Assert.Equal(5, result.Count);
            Assert.Equal(date, result.First().Start);
        }
    }
}
