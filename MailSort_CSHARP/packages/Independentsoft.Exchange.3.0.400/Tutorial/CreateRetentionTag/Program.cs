using System;
using System.Net;
using System.Collections.Generic;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            Service service = new Service("https://myserver/ews/Exchange.asmx", credential);

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

            try
            {
                RetentionSettings settings = service.GetRetentionSettings(StandardFolder.Inbox);

                RetentionPolicyTag myRetentionPolicy = new RetentionPolicyTag();

                myRetentionPolicy.Id = Guid.NewGuid().ToString();
                myRetentionPolicy.ObjectGuid = Guid.NewGuid().ToString();
                myRetentionPolicy.DisplayName = "Delete after 3 months";
                myRetentionPolicy.RetentionAction = RetentionAction.DeleteAndAllowRecovery;
                myRetentionPolicy.Period = 90; //90 days
                myRetentionPolicy.Type = ElcFolderType.Personal;
                myRetentionPolicy.IsArchive = false;
                myRetentionPolicy.IsVisible = true;
                myRetentionPolicy.OptedInto = true;

                myRetentionPolicy.ContentSettings = new ContentSettings();
                myRetentionPolicy.ContentSettings.Id = Guid.NewGuid().ToString();
                myRetentionPolicy.ContentSettings.MessageClass = "*";
                myRetentionPolicy.ContentSettings.Period = 90;
                myRetentionPolicy.ContentSettings.RetentionAction = RetentionAction.DeleteAndAllowRecovery;

                settings.RetentionPolicyTags.Add(myRetentionPolicy);

                Response response = service.UpdateRetentionSettings(settings, StandardFolder.Inbox);
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Error: " + ex.XmlMessage);
                Console.Read();
            }
            catch (WebException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.Read();
            }
        }
    }
}
