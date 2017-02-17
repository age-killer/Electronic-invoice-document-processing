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

                Dim response As GetServerTimeZonesResponse = service.GetServerTimeZones()

                ''Display all available time zones
                For Each timeZone As TimeZoneDefinition In response.TimeZoneDefinitions
                    Console.WriteLine(timeZone.Id)
                    Console.WriteLine(timeZone.Name)
                    Console.WriteLine("-------------------------------------------------------")
                Next

                ''Display time zones with offset -05:00
                For Each timeZone As TimeZoneDefinition In response.TimeZoneDefinitions
                    If (timeZone.Name.IndexOf("UTC-05:00") > 0) Then
                        Console.WriteLine(timeZone.Id)
                        Console.WriteLine(timeZone.Name)
                        Console.WriteLine("-------------------------------------------------------")
                    End If
                Next

                ''Find time zone based on name
                For Each timeZone As TimeZoneDefinition In response.TimeZoneDefinitions
                    If (timeZone.Name.IndexOf("Dublin") > 0) Then
                        Console.WriteLine(timeZone.Id)
                        Console.WriteLine(timeZone.Name)
                        Console.WriteLine("-------------------------------------------------------")
                    End If
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