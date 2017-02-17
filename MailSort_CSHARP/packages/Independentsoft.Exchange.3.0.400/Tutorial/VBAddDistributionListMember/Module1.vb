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

                Dim list As DistributionList = service.GetDistributionList("AAMkADZkZDU1ZTQyLTkZ8lTi+aWAAAAAAAZAABRQo1uip9ARZANZ8lTi+aWAACof8EsAAA=")

                Dim itemId As ItemId = list.ItemId

                Dim memebersProperty As [Property] = New [Property](DistributionListPropertyPath.Members, New DistributionListMember(New Mailbox("john@mydomain.com")))

                Dim itemChange As ItemChange = New ItemChange(itemId)
                itemChange.PropertiesToAppend.Add(memebersProperty)

                itemId = service.UpdateItem(itemChange)

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