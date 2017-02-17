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
                IList<MailboxData> mailboxes = new List<MailboxData>();
                mailboxes.Add(new MailboxData("John@mydomain.com"));
                mailboxes.Add(new MailboxData("Peter@mydomain.com"));
                mailboxes.Add(new MailboxData("Mark@mydomain.com"));

                SerializableTimeZone timeZone = new SerializableTimeZone(StandardTimeZone.Berlin);
                SuggestionsViewOptions suggestionsViewOptions = new SuggestionsViewOptions(DateTime.Today.AddDays(1), DateTime.Today.AddDays(2), 60);

                AvailabilityResponse response = service.GetAvailability(mailboxes, timeZone, suggestionsViewOptions);

                if (response.SuggestionsResponse != null)
                {
                    IList<SuggestionDay> suggestionDays = response.SuggestionsResponse.SuggestionDays;

                    for (int i = 0; i < suggestionDays.Count; i++)
                    {
                        Console.WriteLine("Suggestion Day = " + suggestionDays[i].Date);

                        for (int j = 0; j < suggestionDays[i].Suggestions.Count; j++)
                        {
                            Suggestion suggestion = suggestionDays[i].Suggestions[j];

                            Console.WriteLine("MeetingTime = " + suggestion.MeetingTime);
                            Console.WriteLine("Quality = " + suggestion.Quality);
                        }

                        Console.WriteLine("---------------------------------------------------------");
                    }
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
