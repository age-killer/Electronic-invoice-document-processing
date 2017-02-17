Imports System
Imports System.Net
Imports Independentsoft.Exchange
Imports Independentsoft.Exchange.Autodiscover

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim autodiscoverService As New AutodiscoverService("https://myserver/autodiscover/autodiscover.xml", credential)

            Try

                Dim settingNames As New List(Of UserSettingName)()
                settingNames.Add(UserSettingName.CasVersion)
                settingNames.Add(UserSettingName.EwsSupportedSchemas)
                settingNames.Add(UserSettingName.ExternalEwsUrl)
                settingNames.Add(UserSettingName.ExternalMailboxServer)
                settingNames.Add(UserSettingName.PublicFolderServer)
                settingNames.Add(UserSettingName.UserDeploymentId)
                settingNames.Add(UserSettingName.UserDisplayName)
                settingNames.Add(UserSettingName.UserDN)

                Dim response As GetUserSettingsResponse = autodiscoverService.GetUserSettings("John@mydomain.com", settingNames)

                If response.UserResponses.Count > 0 Then
                    Dim userResponse As UserResponse = response.UserResponses(0)

                    Dim casVersionUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.CasVersion)

                    If casVersionUserSetting IsNot Nothing Then
                        Console.WriteLine("CasVersion=" & casVersionUserSetting.Value)
                    End If

                    Dim ewsSupportedSchemasUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.EwsSupportedSchemas)

                    If ewsSupportedSchemasUserSetting IsNot Nothing Then
                        Console.WriteLine("EwsSupportedSchemas=" & ewsSupportedSchemasUserSetting.Value)
                    End If

                    Dim externalEwsUrlUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.ExternalEwsUrl)

                    If externalEwsUrlUserSetting IsNot Nothing Then
                        Console.WriteLine("ExternalEwsUrl=" & externalEwsUrlUserSetting.Value)
                    End If

                    Dim externalMailboxServerUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.ExternalMailboxServer)

                    If externalMailboxServerUserSetting IsNot Nothing Then
                        Console.WriteLine("ExternalMailboxServer=" & externalMailboxServerUserSetting.Value)
                    End If

                    Dim publicFolderServerUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.PublicFolderServer)

                    If publicFolderServerUserSetting IsNot Nothing Then
                        Console.WriteLine("PublicFolderServer=" & publicFolderServerUserSetting.Value)
                    End If

                    Dim userDeploymentIdUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.UserDeploymentId)

                    If userDeploymentIdUserSetting IsNot Nothing Then
                        Console.WriteLine("UserDeploymentId=" & userDeploymentIdUserSetting.Value)
                    End If

                    Dim userDisplayNameUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.UserDisplayName)

                    If userDisplayNameUserSetting IsNot Nothing Then
                        Console.WriteLine("UserDisplayName=" & userDisplayNameUserSetting.Value)
                    End If

                    Dim userDNUserSetting As UserSetting = userResponse.UserSettings(UserSettingName.UserDN)

                    If userDNUserSetting IsNot Nothing Then
                        Console.WriteLine("UserDN=" & userDNUserSetting.Value)
                    End If
                End If

                Console.Read()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " + ex.Message)
                Console.WriteLine("Error: " + ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " + ex.Message)
                Console.Read()
            End Try

        End Sub
    End Class
End Namespace