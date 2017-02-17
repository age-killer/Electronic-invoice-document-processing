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
                Dim propertyPaths As IList(Of PropertyPath) = New List(Of PropertyPath)()
                propertyPaths.Add(MessagePropertyPath.Subject)
                propertyPaths.Add(MessagePropertyPath.ReceivedTime)
                propertyPaths.Add(MessagePropertyPath.SentTime)
                propertyPaths.Add(MessagePropertyPath.HasAttachments)
                propertyPaths.Add(MessagePropertyPath.IsRead)
                propertyPaths.Add(MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS)
                propertyPaths.Add(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS)
                propertyPaths.Add(MapiPropertyTag.PR_BODY)
                propertyPaths.Add(MapiPropertyTag.PR_HTML)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, propertyPaths)

                For i As Integer = 0 To response.Items.Count - 1
                    If TypeOf response.Items(i) Is Message Then
                        Dim message As Message = DirectCast(response.Items(i), Message)

                        Console.WriteLine("Subject = " & message.Subject)
                        Console.WriteLine("ReceivedTime = " & message.ReceivedTime)
                        Console.WriteLine("SentTime = " & message.SentTime)
                        Console.WriteLine("HasAttachments = " & message.HasAttachments)
                        Console.WriteLine("IsRead = " & message.IsRead)

                        If message.ExtendedProperties(MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS) IsNot Nothing Then
                            Console.WriteLine("From email = " & message.ExtendedProperties(MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS).Value)
                        End If

                        If message.ExtendedProperties(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS) IsNot Nothing Then
                            Console.WriteLine("To = " & message.ExtendedProperties(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS).Value)
                        End If

                        If message.ExtendedProperties(MapiPropertyTag.PR_BODY) IsNot Nothing Then
                            Console.WriteLine("Plain body = " & message.ExtendedProperties(MapiPropertyTag.PR_BODY).Value)
                        End If

                        If message.ExtendedProperties(MapiPropertyTag.PR_HTML) IsNot Nothing Then
                            Dim base64 As Byte() = Convert.FromBase64String(message.ExtendedProperties(MapiPropertyTag.PR_HTML).Value)
                            Dim htmlBody As String = System.Text.Encoding.UTF8.GetString(base64)
                            Console.WriteLine("Html body = " & htmlBody)
                        End If

                        Console.WriteLine("--------------------------------------------------------")
                    End If

                Next

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
    End Class
End Namespace