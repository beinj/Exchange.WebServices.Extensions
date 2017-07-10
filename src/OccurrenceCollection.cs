using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static Microsoft.Exchange.WebServices.Data.Recurrence;

namespace Exchange.WebServices.Extensions
{
    public class OccurrenceCollection : List<Occurrence>
    {
        internal OccurrenceCollection()
        {
        }

        public static async Task<OccurrenceCollection> Bind(ExchangeService service, ItemId recurringMasterId)
        {
            var definition = new PropertyDefinition[]
            {
                AppointmentSchema.Recurrence,
                AppointmentSchema.ModifiedOccurrences,
                AppointmentSchema.DeletedOccurrences
            };

            var propertySet = new PropertySet(BasePropertySet.FirstClassProperties, definition);

            var appointment = await Appointment.Bind(service, recurringMasterId, propertySet);

            if (!appointment.Recurrence.HasEnd)
            {
                ////throw new NotSupportedException("Appointment must has a fixed number of occurrences or an end date.");
                return new OccurrenceCollection();
            }

            var defaultOccurrence = new Occurrence
            {
                Start = appointment.Start,
                End = appointment.End,
                IsAllDayEvent = appointment.IsAllDayEvent,
                IsCancelled = appointment.IsCancelled,
                LastModifiedTime = appointment.LastModifiedTime,
                MasterAppointmentId = appointment.Id.UniqueId,
                Sensitivity = appointment.Sensitivity,
                Subject = appointment.Subject,
                Text = appointment.TextBody?.Text
            };

            var result = new OccurrenceCollection();

            if (appointment.Recurrence is DailyPattern dailyPattern)
            {
                return PatternConverter.Convert(dailyPattern, defaultOccurrence);
            }
            else if (appointment.Recurrence is WeeklyPattern weeklyPattern)
            {
                return PatternConverter.Convert(weeklyPattern, defaultOccurrence);
            }
            else if (appointment.Recurrence is MonthlyPattern monthlyPattern)
            {
                return PatternConverter.Convert(monthlyPattern, defaultOccurrence);
            }
            else if (appointment.Recurrence is YearlyPattern yearlyPattern)
            {
                return PatternConverter.Convert(yearlyPattern, defaultOccurrence);
            }
            else if (appointment.Recurrence is RelativeMonthlyPattern relativeMonthlyPattern)
            {
                return PatternConverter.Convert(relativeMonthlyPattern, defaultOccurrence);
            }
            else if (appointment.Recurrence is RelativeYearlyPattern relativeYearlyPattern)
            {
                return PatternConverter.Convert(relativeYearlyPattern, defaultOccurrence);
            }

            throw new NotSupportedException();
        }
    }
}
