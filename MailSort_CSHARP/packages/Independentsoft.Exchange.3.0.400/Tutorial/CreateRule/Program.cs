using System;
using System.Net;
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
                Rule rule = new Rule();
                rule.DisplayName = "TestRule";

                rule.Conditions = new RuleConditions();
                rule.Conditions.ContainsSubjectStrings.Add("REMOVEME");

                rule.Actions = new RuleActions();
                rule.Actions.Delete = true;

                RuleOperationResponse response = service.CreateRule(rule);
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
