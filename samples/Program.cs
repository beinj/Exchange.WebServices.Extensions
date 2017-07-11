using Exchange.WebServices.Extensions;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using System;

namespace samples
{
    public static class Program
    {
        static void Main(string[] args)
        {
            Run().Wait();
        }

        private static async System.Threading.Tasks.Task Run()
        {
            var configuration = new ConfigurationBuilder()
                .AddUserSecrets("a8aa09bf-0683-4154-81b4-9fd6392e0a74")
                .Build();

            var webCredentials = new WebCredentials(configuration["Username"], configuration["Password"]);
            var impersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, configuration["Email"]);
            var uri = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            var service = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
            {
                ImpersonatedUserId = impersonatedUserId,
                Credentials = webCredentials,
                Url = uri
            };

            var folderId = new FolderId(WellKnownFolderName.Calendar);
            var propertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.Start, AppointmentSchema.AppointmentType);

            var syncFolder = await service.SyncFolderItems(folderId, propertySet, null, 512, SyncFolderItemsScope.NormalItems, null);

            foreach (var itemChange in syncFolder)
            {
                if (itemChange.Item is Appointment appointment && appointment.AppointmentType == AppointmentType.RecurringMaster)
                {
                    var collection = await OccurrenceCollection.Bind(service, appointment.Id);
                    foreach (var item in collection)
                    {
                        Console.WriteLine(item.Start.ToString("g"));
                    }
                }
            }
        }
    }
}
