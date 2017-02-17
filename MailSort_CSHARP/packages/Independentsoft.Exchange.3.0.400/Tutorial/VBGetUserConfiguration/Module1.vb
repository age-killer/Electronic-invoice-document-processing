Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim configName As New UserConfigurationName("CategoryList", New StandardFolderId(StandardFolder.Calendar))

                Dim response As GetUserConfigurationResponse = service.GetUserConfiguration(configName, UserConfigurationProperty.All)

                If response.UserConfiguration IsNot Nothing AndAlso response.UserConfiguration.XmlData IsNot Nothing Then
                    Dim xmlDataBuffer As Byte() = Convert.FromBase64String(response.UserConfiguration.XmlData)

                    Dim xmlData As String = System.Text.Encoding.UTF8.GetString(xmlDataBuffer)

                    Console.WriteLine(xmlData)
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