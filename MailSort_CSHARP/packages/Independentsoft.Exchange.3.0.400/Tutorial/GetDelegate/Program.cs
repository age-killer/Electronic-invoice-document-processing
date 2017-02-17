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

            try
            {
                Mailbox mailbox = new Mailbox("John@mydomain.com");

                DelegateResponse response = service.GetDelegate(mailbox, true);

                IList<DelegateUserResponse> delegateUserResponse = response.DelegateUserResponses;

                for (int i = 0; i < delegateUserResponse.Count; i++)
                {
                    DelegateUser user = delegateUserResponse[i].DelegateUser;

                    Console.WriteLine("User = " + user.UserId);
                    Console.WriteLine("CalendarFolderPermissionLevel = " + user.CalendarFolderPermissionLevel);
                    Console.WriteLine("ContactsFolderPermissionLevel = " + user.ContactsFolderPermissionLevel);
                    Console.WriteLine("InboxFolderPermissionLevel = " + user.InboxFolderPermissionLevel);
                    Console.WriteLine("JournalFolderPermissionLevel = " + user.JournalFolderPermissionLevel);
                    Console.WriteLine("NotesFolderPermissionLevel = " + user.NotesFolderPermissionLevel);
                    Console.WriteLine("ReceiveCopiesOfMeetingMessages = " + user.ReceiveCopiesOfMeetingMessages);
                    Console.WriteLine("TasksFolderPermissionLevel = " + user.TasksFolderPermissionLevel);
                    Console.WriteLine("ViewPrivateItems = " + user.ViewPrivateItems);
                    Console.WriteLine("------------------------------------------------------------------------");
                }

                Console.Read();
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
