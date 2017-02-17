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
                UserConfigurationDictionaryKey key1 = new UserConfigurationDictionaryKey(UserConfigurationDictionaryObjectType.String, "key1");
                UserConfigurationDictionaryValue value1 = new UserConfigurationDictionaryValue(UserConfigurationDictionaryObjectType.String, "value1");

                UserConfigurationDictionaryKey key2 = new UserConfigurationDictionaryKey(UserConfigurationDictionaryObjectType.String, "key2");
                UserConfigurationDictionaryValue value2 = new UserConfigurationDictionaryValue(UserConfigurationDictionaryObjectType.String, "value2");

                UserConfigurationDictionaryEntry entry1 = new UserConfigurationDictionaryEntry(key1, value1);
                UserConfigurationDictionaryEntry entry2 = new UserConfigurationDictionaryEntry(key2, value2);

                UserConfiguration userConfiguration = new UserConfiguration();
                userConfiguration.Name = new UserConfigurationName("config1", new StandardFolderId(StandardFolder.Drafts));
                userConfiguration.Entries.Add(entry1);
                userConfiguration.Entries.Add(entry2);

                Response response = service.CreateUserConfiguration(userConfiguration);
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
