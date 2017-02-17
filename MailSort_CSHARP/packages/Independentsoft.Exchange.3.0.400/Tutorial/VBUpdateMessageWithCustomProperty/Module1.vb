Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New Contains(MessagePropertyPath.Subject, "test", ContainmentMode.Prefixed, ContainmentComparison.IgnoreCase)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim myPropertyName As New PropertyName("MyCustomProperty", StandardPropertySet.PublicStrings, MapiPropertyType.[String])
                        Dim myProperty As New ExtendedProperty(myPropertyName, "value1")

                        Dim newItemId As ItemId = service.UpdateItem(response.Items(i).ItemId, myProperty)

                    End If
                Next

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