Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim settings As RetentionSettings = service.GetRetentionSettings(StandardFolder.Inbox)

                For Each tag As RetentionPolicyTag In settings.RetentionPolicyTags

                    Console.WriteLine(tag.Id)
                    Console.WriteLine(tag.DisplayName)
                    Console.WriteLine(tag.RetentionAction.ToString())
                    Console.WriteLine(tag.Period)
                    Console.WriteLine(tag.IsArchive)
                    Console.WriteLine(tag.IsVisible)
                    Console.WriteLine("---------------------------------------------")

                Next

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