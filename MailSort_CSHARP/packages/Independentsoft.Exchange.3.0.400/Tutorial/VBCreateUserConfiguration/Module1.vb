Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim key1 As New UserConfigurationDictionaryKey(UserConfigurationDictionaryObjectType.[String], "key1")
                Dim value1 As New UserConfigurationDictionaryValue(UserConfigurationDictionaryObjectType.[String], "value1")

                Dim key2 As New UserConfigurationDictionaryKey(UserConfigurationDictionaryObjectType.[String], "key2")
                Dim value2 As New UserConfigurationDictionaryValue(UserConfigurationDictionaryObjectType.[String], "value2")

                Dim entry1 As New UserConfigurationDictionaryEntry(key1, value1)
                Dim entry2 As New UserConfigurationDictionaryEntry(key2, value2)

                Dim userConfiguration As New UserConfiguration()
                userConfiguration.Name = New UserConfigurationName("config1", New StandardFolderId(StandardFolder.Drafts))
                userConfiguration.Entries.Add(entry1)
                userConfiguration.Entries.Add(entry2)

                Dim response As Response = service.CreateUserConfiguration(userConfiguration)

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