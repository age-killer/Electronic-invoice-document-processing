Imports System
Imports System.Net
Imports System.Collections
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1

        Private Shared folderTable As New Hashtable()

        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                Dim offset As Integer = 0

                While (True)

                    Dim view As IndexedPageView = New IndexedPageView(offset, IndexBasePoint.Beginning, 100)

                    Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, view)

                    For i As Integer = 0 To response.Items.Count - 1

                        If TypeOf response.Items(i) Is Message Then

                            Dim message As Message = DirectCast(response.Items(i), Message)

                            Console.WriteLine("Subject = " + message.Subject)
                            Console.WriteLine("ReceivedTime = " + message.ReceivedTime)
                            Console.WriteLine("----------------------------------------------------------------")
                        End If
                    Next

                    If (response.IndexedPagingOffset < response.TotalItemsInView) Then
                        offset = response.IndexedPagingOffset
                    Else

                        Exit While
                    End If

                End While

                Console.Read()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try
        End Sub

        Private Shared Function GetFolderPath(ByVal folder As Folder) As String
            Dim path As String = ""

            path = (path & "/") + folder.DisplayName

            Dim parentId As String = folder.ParentId.Id

            If folderTable(parentId) IsNot Nothing Then
                Dim parentFolder As Folder = DirectCast(folderTable(parentId), Folder)

                Return GetFolderPath(parentFolder) + path
            End If

            Return path
        End Function

    End Class
End Namespace
