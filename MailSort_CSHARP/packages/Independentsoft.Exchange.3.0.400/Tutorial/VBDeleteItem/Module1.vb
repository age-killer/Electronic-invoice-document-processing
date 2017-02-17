Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Delete Draft items older then 1 month 

                Dim restriction As New IsLessThan(ItemPropertyPath.CreatedTime, DateTime.Today.AddMonths(-1))

                Dim draftsItems As FindItemResponse = service.FindItem(StandardFolder.Drafts, restriction)

                For i As Integer = 0 To draftsItems.Items.Count - 1

                    Dim response As Response = service.DeleteItem(draftsItems.Items(i).ItemId, DeleteType.MoveToDeletedItems)
                Next

                'Empty DeletedItems folder 

                Dim itemShape As New ItemShape(ShapeType.Id)
                Dim deletedItems As FindItemResponse = service.FindItem(StandardFolder.DeletedItems, itemShape)

                Dim itemIds As New List(Of ItemId)()

                For i As Integer = 0 To deletedItems.Items.Count - 1

                    itemIds.Add(deletedItems.Items(i).ItemId)
                Next

                Dim responses As IList(Of Response) = service.DeleteItem(itemIds, DeleteType.HardDelete)

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