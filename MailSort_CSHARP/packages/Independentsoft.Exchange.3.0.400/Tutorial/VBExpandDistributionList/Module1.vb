Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(ContactPropertyPath.ItemClass, ItemClass.DistributionList)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Contacts, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    Dim itemId As ItemId = response.Items(i).ItemId

                    Dim listResponse As ExpandDistributionListResponse = service.ExpandDistributionList(itemId)
                    Dim mailboxes As IList(Of Mailbox) = listResponse.Mailboxes

                    For j As Integer = 0 To mailboxes.Count - 1
                        Console.WriteLine("Member = " + mailboxes(j).EmailAddress)
                    Next

                    Console.WriteLine("---------------------------------------------------------------")
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