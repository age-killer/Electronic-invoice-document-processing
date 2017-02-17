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
                IList<UserSettingName> settingNames = new List<UserSettingName>();
                settingNames.Add(UserSettingName.CasVersion);
                settingNames.Add(UserSettingName.EwsSupportedSchemas);
                settingNames.Add(UserSettingName.ExternalEwsUrl);
                settingNames.Add(UserSettingName.ExternalMailboxServer);
                settingNames.Add(UserSettingName.PublicFolderServer);
                settingNames.Add(UserSettingName.UserDeploymentId);
                settingNames.Add(UserSettingName.UserDisplayName);
                settingNames.Add(UserSettingName.UserDN);

                GetUserSettingsResponse response = autodiscoverService.GetUserSettings("John@mydomain.com", settingNames);

                if (response.UserResponses.Count > 0)
                {
                    UserResponse userResponse = response.UserResponses[0];

                    UserSetting casVersionUserSetting = userResponse.UserSettings[UserSettingName.CasVersion];

                    if (casVersionUserSetting != null)
                    {
                        Console.WriteLine("CasVersion=" + casVersionUserSetting.Value);
                    }

                    UserSetting ewsSupportedSchemasUserSetting = userResponse.UserSettings[UserSettingName.EwsSupportedSchemas];

                    if (ewsSupportedSchemasUserSetting != null)
                    {
                        Console.WriteLine("EwsSupportedSchemas=" + ewsSupportedSchemasUserSetting.Value);
                    }

                    UserSetting externalEwsUrlUserSetting = userResponse.UserSettings[UserSettingName.ExternalEwsUrl];

                    if (externalEwsUrlUserSetting != null)
                    {
                        Console.WriteLine("ExternalEwsUrl=" + externalEwsUrlUserSetting.Value);
                    }

                    UserSetting externalMailboxServerUserSetting = userResponse.UserSettings[UserSettingName.ExternalMailboxServer];

                    if (externalMailboxServerUserSetting != null)
                    {
                        Console.WriteLine("ExternalMailboxServer=" + externalMailboxServerUserSetting.Value);
                    }

                    UserSetting publicFolderServerUserSetting = userResponse.UserSettings[UserSettingName.PublicFolderServer];

                    if (publicFolderServerUserSetting != null)
                    {
                        Console.WriteLine("PublicFolderServer=" + publicFolderServerUserSetting.Value);
                    }

                    UserSetting userDeploymentIdUserSetting = userResponse.UserSettings[UserSettingName.UserDeploymentId];

                    if (userDeploymentIdUserSetting != null)
                    {
                        Console.WriteLine("UserDeploymentId=" + userDeploymentIdUserSetting.Value);
                    }

                    UserSetting userDisplayNameUserSetting = userResponse.UserSettings[UserSettingName.UserDisplayName];

                    if (userDisplayNameUserSetting != null)
                    {
                        Console.WriteLine("UserDisplayName=" + userDisplayNameUserSetting.Value);
                    }

                    UserSetting userDNUserSetting = userResponse.UserSettings[UserSettingName.UserDN];

                    if (userDNUserSetting != null)
                    {
                        Console.WriteLine("UserDN=" + userDNUserSetting.Value);
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
