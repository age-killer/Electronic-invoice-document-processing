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

            try
            {
                UserConfigurationName configName = new UserConfigurationName("CategoryList", new StandardFolderId(StandardFolder.Calendar));

                GetUserConfigurationResponse response = service.GetUserConfiguration(configName, UserConfigurationProperty.All);

                if (response.UserConfiguration != null && response.UserConfiguration.XmlData != null)
                {
                    byte[] xmlDataBuffer = Convert.FromBase64String(response.UserConfiguration.XmlData);

                    string xmlData = System.Text.Encoding.UTF8.GetString(xmlDataBuffer);

                    Console.WriteLine(xmlData);
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
