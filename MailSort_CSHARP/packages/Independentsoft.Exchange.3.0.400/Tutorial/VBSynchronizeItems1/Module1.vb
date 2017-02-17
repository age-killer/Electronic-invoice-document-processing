Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Initial synscronization. It will retrieve all existing items from the specified folder 
                Dim response As SyncItemsResponse = service.SyncItems(StandardFolder.Inbox)

                'Keep synchronization state in order to use it in next synchronizations 
                Dim state As String = response.State

                Dim allExistingItems As IList(Of Item) = response.CreatedItems

                Console.WriteLine("Wait until new messages arrive or create/mark read/delete messages in the Inbox folder and press ENTER")
                Console.Read()

                'Synchronize items again but now include the state parameter in order to find changes from last synchronization 
                Dim newResponse As SyncItemsResponse = service.SyncItems(StandardFolder.Inbox, state)

                Dim newItems As IList(Of Item) = newResponse.CreatedItems
                Dim deletedItems As IList(Of ItemId) = newResponse.DeletedItems
                Dim updatedItems As IList(Of Item) = newResponse.UpdatedItems
                Dim readFlagChangedItems As IList(Of ReadFlagChange) = newResponse.ReadFlagChangedItems

                'Display new items 
                For i As Integer = 0 To newItems.Count - 1

                    Console.WriteLine("New = " & newItems(i).Subject)
                    Console.WriteLine("New ItemId= " & newItems(i).ItemId.ToString())
                Next

                'Display deleted items 
                For i As Integer = 0 To deletedItems.Count - 1

                    Console.WriteLine("Deleted ItemId = " & deletedItems(i).ToString())
                Next

                'Display updated items 
                For i As Integer = 0 To updatedItems.Count - 1

                    Console.WriteLine("Updated = " & updatedItems(i).Subject)
                    Console.WriteLine("Updated ItemId = " & updatedItems(i).ItemId.ToString())
                Next

                'Display items marked read/unread 
                For i As Integer = 0 To readFlagChangedItems.Count - 1

                    Console.WriteLine("Marked read/unread = " & readFlagChangedItems(i).ItemId.ToString())
                    Console.WriteLine("IsRead = " & readFlagChangedItems(i).IsRead)
                Next

                Console.Read()
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