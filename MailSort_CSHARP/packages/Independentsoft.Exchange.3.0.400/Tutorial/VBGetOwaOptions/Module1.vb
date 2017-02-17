Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try

                Dim options As OwaOptions = service.GetOwaOptions()

                Console.WriteLine("AlwaysShowBcc:" & options.AlwaysShowBcc)
                Console.WriteLine("AlwaysShowFrom:" & options.AlwaysShowFrom)
                Console.WriteLine("AutoAddSignature:" & options.AutoAddSignature)
                Console.WriteLine("CheckNameInContactsFirst:" & options.CheckNameInContactsFirst)
                Console.WriteLine("ComposeMarkup:" & options.ComposeMarkup)
                Console.WriteLine("ConversationSortOrder:" & options.ConversationSortOrder)
                Console.WriteLine("DateFormat:" & options.DateFormat)
                Console.WriteLine("EmptyDeletedItemsOnLogoff:" & options.EmptyDeletedItemsOnLogoff)
                Console.WriteLine("EnableReminders:" & options.EnableReminders)
                Console.WriteLine("FirstWeekOfYear:" & options.FirstWeekOfYear)
                Console.WriteLine("FormatBarState:" & options.FormatBarState)
                Console.WriteLine("HideDeletedItems:" & options.HideDeletedItems)
                Console.WriteLine("HourIncrement:" & options.HourIncrement)
                Console.WriteLine("IsOptimizedForAccessibility:" & options.IsOptimizedForAccessibility)
                Console.WriteLine("IsQuicklinksBarVisible:" & options.IsQuicklinksBarVisible)
                Console.WriteLine("MarkAsReadDelayTime:" & options.MarkAsReadDelayTime)
                Console.WriteLine("NewItemNotify:" & options.NewItemNotify)
                Console.WriteLine("NextSelection:" & options.NextSelection)
                Console.WriteLine("PreviewMarkAsRead:" & options.PreviewMarkAsRead)
                Console.WriteLine("ReadReceipt:" & options.ReadReceipt)
                Console.WriteLine("SendAsItemsCopiedTo:" & options.SendAsItemsCopiedTo)
                Console.WriteLine("SendOnBehalfOfCopiedTo:" & options.SendOnBehalfOfCopiedTo)
                Console.WriteLine("ShowTreeInListView:" & options.ShowTreeInListView)
                Console.WriteLine("ShowWeekNumbers:" & options.ShowWeekNumbers)
                Console.WriteLine("SignatureHtml:" & options.SignatureHtml)
                Console.WriteLine("SignatureText:" & options.SignatureText)
                Console.WriteLine("SpellingCheckBeforeSend:" & options.SpellingCheckBeforeSend)
                Console.WriteLine("SpellingDictionaryLanguage:" & options.SpellingDictionaryLanguage)
                Console.WriteLine("SpellingIgnoreMixedDigits:" & options.SpellingIgnoreMixedDigits)
                Console.WriteLine("SpellingIgnoreUpperCase:" & options.SpellingIgnoreUpperCase)
                Console.WriteLine("ThemeStorageId:" & options.ThemeStorageId)
                Console.WriteLine("TimeFormat:" & options.TimeFormat)
                Console.WriteLine("TimeZone:" & options.TimeZone)
                Console.WriteLine("WeekStartDay:" & options.WeekStartDay)

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