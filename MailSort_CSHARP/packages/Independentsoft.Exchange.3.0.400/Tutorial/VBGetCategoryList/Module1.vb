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

                Dim list As CategoryList = service.GetCategoryList()

                For Each category As Category In List.Categories

                    Console.WriteLine(category.Name)
                    Console.WriteLine(category.Color.ToString)
                    Console.WriteLine("--------------------------------------------")
                Next

                ''if you want to delete category from the list
                For i As Integer = list.Categories.Count - 1 To i > 0 Step -1

                    If list.Categories(i).Name = "MyCategory" Then
                        list.Categories.RemoveAt(i)
                    End If
                Next

                ''save changes after delete
                service.UpdateCategoryList(List)

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