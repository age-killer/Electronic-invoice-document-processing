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
                IsEqualTo restriction = new IsEqualTo(NotePropertyPath.ItemClass, ItemClass.Note);

                FindItemResponse response = service.FindItem(StandardFolder.Notes, NotePropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    Note note = (Note)response.Items[i];

                    Console.WriteLine("Subject = " + note.Subject);
                    Console.WriteLine("Width = " + note.Width);
                    Console.WriteLine("Height = " + note.Height);
                    Console.WriteLine("Left = " + note.Left);
                    Console.WriteLine("Top = " + note.Top);
                    Console.WriteLine("Color = " + note.Color);
                    Console.WriteLine("IconColor = " + note.IconColor);
                    Console.WriteLine("Body Preview = " + note.BodyPlainText);
                    Console.WriteLine("----------------------------------------------------------------");
                }

                Console.Read();
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
