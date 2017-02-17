Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim member1 As New DistributionListMember(New Mailbox("John@mydomain.com"))
                Dim member2 As New DistributionListMember(New Mailbox("Mark@mydomain.com"))

                Dim list As New DistributionList()
                list.DisplayName = "Test"
                list.Subject = "Test"
                list.Body = New Body("Body text")
                list.Members.Add(member1)
                list.Members.Add(member2)

                Dim itemId As ItemId = service.CreateItem(list)

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