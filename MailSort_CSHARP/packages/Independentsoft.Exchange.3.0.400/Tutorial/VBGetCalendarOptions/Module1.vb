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

                Dim options As CalendarOptions = service.GetCalendarOptions()

                If options IsNot Nothing Then

                    Console.WriteLine("AutoProcess:" & options.AutoProcess)
                    Console.WriteLine("AutoProcessExternal:" & options.AutoProcessExternal)
                    Console.WriteLine("AutoProcessNewItems:" & options.AutoProcessNewItems)
                    Console.WriteLine("DeleteUpdatedItems:" & options.DeleteUpdatedItems)
                    Console.WriteLine("ReminderDefault:" & options.ReminderDefault)
                    Console.WriteLine("ReminderUpgradeTime:" & options.ReminderUpgradeTime)
                    Console.WriteLine("RemoveForwardedMeetingNotifications:" & options.RemoveForwardedMeetingNotifications)

                End If

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