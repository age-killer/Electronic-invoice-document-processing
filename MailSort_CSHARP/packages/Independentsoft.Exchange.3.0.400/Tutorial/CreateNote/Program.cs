using System;
using System.Net;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            Service service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                Note note = new Note();
                note.Subject = "My test note";
                note.Body = new Body("My test note");
                note.Color = NoteColor.Green;
                note.IconColor = NoteColor.Green;
                note.Height = 200;
                note.Width = 300;
                note.Left = 400;
                note.Top = 200;

                ItemId itemId = service.CreateItem(note, StandardFolder.Notes);
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Error: " + ex.XmlMessage);
                Console.Read();
            }
            catch (WebException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.Read();
            }
        }
    }
}
