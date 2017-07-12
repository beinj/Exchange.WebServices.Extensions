using Microsoft.Exchange.WebServices.Data;
using System;

namespace Exchange.WebServices.OccurrenceData
{
    public class Occurrence
    {
        public Occurrence()
        {

        }

        internal Occurrence(Occurrence appointment)
        {
            IsAllDayEvent = appointment.IsAllDayEvent;
            IsCancelled = appointment.IsCancelled;
            LastModifiedTime = appointment.LastModifiedTime;
            MasterAppointmentId = appointment.MasterAppointmentId;
            Sensitivity = appointment.Sensitivity;
            Subject = appointment.Subject;
            Text = appointment.Text;
        }

        public int Index { get; set; }

        public string MasterAppointmentId { get; set; }

        public DateTime Start { get; set; }

        public DateTime End { get; set; }

        public bool IsCancelled { get; set; }

        public bool IsAllDayEvent { get; set; }

        public string Subject { get; set; }

        public string Text { get; set; }

        public Sensitivity Sensitivity { get; set; }

        public DateTime LastModifiedTime { get; set; }
    }
}
