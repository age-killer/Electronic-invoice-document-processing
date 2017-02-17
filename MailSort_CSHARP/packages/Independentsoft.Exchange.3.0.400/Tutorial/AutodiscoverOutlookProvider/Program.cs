using System;
using System.Net;
using System.Collections.Generic;
using Independentsoft.Exchange;
using Independentsoft.Exchange.Autodiscover;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            AutodiscoverService autodiscoverService = new AutodiscoverService("https://myserver/autodiscover/autodiscover.xml", credential);

            try
            {
                OutlookProvider outlookProvider = autodiscoverService.AutodiscoverOutlookProvider("John@mydomain.com");

                if (outlookProvider.User != null)
                {
                    User user = outlookProvider.User;

                    Console.WriteLine("DisplayName = " + user.DisplayName);
                    Console.WriteLine("DeploymentId = " + user.DeploymentId);
                    Console.WriteLine("LegacyDN = " + user.LegacyDN);
                }

                if (outlookProvider.Account != null)
                {
                    Account account = outlookProvider.Account;

                    ExchangeProtocol exchangeProtocol = account.ExchangeProtocol;
                    WebProtocol webProtocol = account.WebProtocol;

                    if (exchangeProtocol != null)
                    {
                        Console.WriteLine("Server = " + exchangeProtocol.Server);
                        Console.WriteLine("ServerDN = " + exchangeProtocol.ServerDN);
                        Console.WriteLine("ServerVersion = " + exchangeProtocol.ServerVersion);
                        Console.WriteLine("ActiveDirectory = " + exchangeProtocol.ActiveDirectory);
                        Console.WriteLine("AuthenticationPackage = " + exchangeProtocol.AuthenticationPackage);
                        Console.WriteLine("MailboxDatabaseLegacyDN = " + exchangeProtocol.MailboxDatabaseLegacyDN);
                        Console.WriteLine("AvailabilityServiceUrl = " + exchangeProtocol.AvailabilityServiceUrl);
                        Console.WriteLine("ExchangeWebServiceUrl = " + exchangeProtocol.ExchangeWebServiceUrl);
                        Console.WriteLine("OfflineAddressBookUrl = " + exchangeProtocol.OfflineAddressBookUrl);
                        Console.WriteLine("OutOfOfficeUrl = " + exchangeProtocol.OutOfOfficeUrl);
                        Console.WriteLine("UnifiedMessagingServiceUrl = " + exchangeProtocol.UnifiedMessagingServiceUrl);
                    }

                    if (webProtocol != null)
                    {
                        if (webProtocol.InternalAccess != null)
                        {
                            IList<OutlookWebAccessUrl> owaUrls = webProtocol.InternalAccess.OutlookWebAccessUrls;

                            for (int i = 0; i < owaUrls.Count; i++)
                            {
                                Console.WriteLine("OWA Url = " + owaUrls[i].Url);

                                for (int j = 0; j < owaUrls[i].AuthenticationMethods.Count; j++)
                                {
                                    Console.WriteLine("OWA Authentication = " + owaUrls[i].AuthenticationMethods[j]);
                                }
                            }
                        }

                        if (webProtocol.ExternalAccess != null)
                        {
                            IList<OutlookWebAccessUrl> owaUrls = webProtocol.ExternalAccess.OutlookWebAccessUrls;

                            for (int i = 0; i < owaUrls.Count; i++)
                            {
                                Console.WriteLine("OWA Url = " + owaUrls[i].Url);

                                for (int j = 0; j < owaUrls[i].AuthenticationMethods.Count; j++)
                                {
                                    Console.WriteLine("OWA Authentication = " + owaUrls[i].AuthenticationMethods[j]);
                                }
                            }
                        }
                    }
                }

                Console.Read();
            }
            catch (AutodiscoverException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
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
