Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New Contains(MessagePropertyPath.Attachments, "test", ContainmentMode.Prefixed, ContainmentComparison.IgnoreCase)

                Dim itemShape As New ItemShape(ShapeType.Id)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, itemShape, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim itemId As ItemId = response.Items(i).ItemId

                        Dim message As Message = service.GetMessage(itemId)

                        Dim attachmentsInfo As IList(Of AttachmentInfo) = message.Attachments

                        For j As Integer = 0 To attachmentsInfo.Count - 1

                            Dim attachment As Attachment = service.GetAttachment(attachmentsInfo(j).Id)

                            If TypeOf attachment Is FileAttachment Then

                                Dim fileAttachment As FileAttachment = DirectCast(attachment, FileAttachment)

                                Dim buffer As Byte() = fileAttachment.Content
                                Dim fileName As String = fileAttachment.Name

                                Dim fileStream As New FileStream("c:\test\" + fileName, FileMode.Create)

                                Using fileStream
                                    fileStream.Write(buffer, 0, buffer.Length)
                                End Using

                            ElseIf TypeOf attachment Is ItemAttachment Then

                                Dim itemAttachment As ItemAttachment = DirectCast(attachment, ItemAttachment)

                                Dim name As String = itemAttachment.Name
                                Dim attachedItem As Item = itemAttachment.Item

                            End If
                        Next
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