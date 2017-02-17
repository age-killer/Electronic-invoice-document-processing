/*
 * Created by SharpDevelop.
 * User: user
 * Date: 2/1/2017
 * Time: 4:26 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
 
using System;
using System.Collections;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Collections.Generic;
using System.Net.Mail;
using System.Configuration;
using System.IO;

namespace MailSort_GIT
{
	class Program
	{
		public static void Main(string[] args)
		{

           Console.WriteLine("Connecting to Exchange");
           
     //    Connection string and save folder
           var address = ConfigurationManager.AppSettings["url"];
     //    var shared_mail = ConfigurationManager.AppSettings["shared_mail"];
           var credentials = ConfigurationManager.AppSettings["credentials"].Split(',');
           var testMail = ConfigurationManager.AppSettings["testMail"];
           var einvoiceMail = ConfigurationManager.AppSettings["einvoiceMail"];
           var saveFolderTesting = ConfigurationManager.AppSettings["saveFolderTesting"];
           var save_Einv = ConfigurationManager.AppSettings["save_Einv"];
                
     //    Arrays with senders' mails
           var cat_SE = ConfigurationManager.AppSettings["se"]; //.Split(',');
           var cat_FU = ConfigurationManager.AppSettings["fu"]; //.Split(',');
           var cat_AS = ConfigurationManager.AppSettings["as"]; //.Split(',');
           var cat_IT = ConfigurationManager.AppSettings["it"]; //.Split(',');
           var cat_MA = ConfigurationManager.AppSettings["ma"]; //.Split(',');
           var cat_OT = ConfigurationManager.AppSettings["ot"]; //.Split(',');
           var cat_SH = ConfigurationManager.AppSettings["sh"]; //.Split(',');
           var cat_UT = ConfigurationManager.AppSettings["ut"]; //.Split(',');
           var cat_UNI = ConfigurationManager.AppSettings["uni"]; //.Split(',');
           var testSender = ConfigurationManager.AppSettings["testSender"]; //.Split(',');
           ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
     
		   service.Credentials = new NetworkCredential(credentials[0], credentials[1], credentials[2]);
           service.Url = new Uri(address);
           Console.WriteLine("Connected, counting unread.");
           Mailbox mailMy = new Mailbox(einvoiceMail);
     //    Mailbox mailMy = new Mailbox(testMail);
     //      Console.WriteLine(mailMy);
           FolderId InboxID = new FolderId(WellKnownFolderName.Inbox, mailMy);
           Console.WriteLine(InboxID);
           
           Folder InboxF = Folder.Bind(service, InboxID);
                  
            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            FindItemsResults<Item> findResults = service.FindItems(InboxID, sf, new ItemView(20));
            Console.WriteLine("Find results are: " + findResults.TotalCount);
            
            foreach (EmailMessage mail in findResults.Items)
            {
            	mail.Load();
            	Console.WriteLine("Entered FOR loop.");
            	         
                String sender = mail.From.Address.ToString();
               // Console.WriteLine(sender);              
                
               // Checking cat_SE
                if (cat_SE.Contains(sender) && mail.HasAttachments)
            	{
                	
                	int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5SE\\" + attNameSplit[0] + ".pdf");
                		   
                		 //   Console.WriteLine("Number SE invoices processed: " + attCounter
                		}
                	}
                	
                  	 File.Create(save_Einv + "e5SE\\" + "invoice.txt").Dispose();
                  	 //var seCount = Directory.EnumerateFiles(save_Einv + "e5SE\\", "*.*", SearchOption.AllDirectories) select file).Count();
                  	 int seCount = new DirectoryInfo(save_Einv + "e5SE\\").GetFiles().Length;
                  	 
                	 Console.WriteLine("Number SE invoices processed: " + seCount);
                //	Console.ReadLine();
            	}
            	
                // Checking cat_FU                 
            	else if (cat_FU.Contains(sender) && mail.HasAttachments)
            	{
            		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5FU\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5FU\\" + "invoice.txt").Dispose();
                		}
                	}
                	
                	 File.Create(save_Einv + "e5FU\\" + "invoice.txt").Dispose();
                  //	 var fuCount = Directory.EnumerateFiles(save_Einv + "e5FU\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int fuCount = new DirectoryInfo(save_Einv + "e5FU\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + fuCount);
            	}
            	
            	// Checking cat_IT  
            	else if (cat_IT.Contains(sender) && mail.HasAttachments)
            	{
            		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5IT\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5IT\\" + "invoice.txt").Dispose();
                		}
                	}
                	
                	 File.Create(save_Einv + "e5IT\\" + "invoice.txt").Dispose();
              //    	 var itCount = Directory.EnumerateFiles(save_Einv + "e5IT\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int itCount = new DirectoryInfo(save_Einv + "e5IT\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + itCount);
            	}
            	
            	// Checking cat_MA
              	else if (cat_MA.Contains(sender) && mail.HasAttachments)
            	{      
            
              	    int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5MA\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5MA\\" + "invoice.txt").Dispose();
                		}
                	}
                	
                	
                	 File.Create(save_Einv + "e5MA\\" + "invoice.txt").Dispose();
                 // 	 var maCount = Directory.EnumerateFiles(save_Einv + "e5MA\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int maCount = new DirectoryInfo(save_Einv + "e5MA\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + maCount);
                }
              	
              	// Checking cat_OT
              	else if (cat_OT.Contains(sender) && mail.HasAttachments)
            	{
              		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5OT\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5OT\\" + "invoice.txt").Dispose();
                		}
                	}
                	
                	                	
                	 File.Create(save_Einv + "e5OT\\" + "invoice.txt").Dispose();
                  //	 var otCount = Directory.EnumerateFiles(save_Einv + "e5OT\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int otCount = new DirectoryInfo(save_Einv + "e5OT\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + otCount);
              	}
              	
              	// Checking cat_SH
              	else if (cat_SH.Contains(sender) && mail.HasAttachments)
            	{
              		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5SH\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5SH\\" + "invoice.txt").Dispose();
                		}
                	}

                	                	
                	 File.Create(save_Einv + "e5SH\\" + "invoice.txt").Dispose();
                 // 	 var shCount = Directory.EnumerateFiles(save_Einv + "e5SH\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int shCount = new DirectoryInfo(save_Einv + "e5SH\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + shCount);                	
              	}
              	
              	// Checking cat_UT
              	else if (cat_UT.Contains(sender) && mail.HasAttachments)
            	{
              		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                		
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "e5UT\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "e5UT\\" + "invoice.txt").Dispose();
                   		}
                	}
                	
                	                	                	
                	 File.Create(save_Einv + "e5UT\\" + "invoice.txt").Dispose();
          //        	 var utCount = Directory.EnumerateFiles(save_Einv + "e5UT\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int utCount = new DirectoryInfo(save_Einv + "e5UT\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + utCount);  
              	}
              	
              	// Checking cat_UNI
              	else if (cat_UNI.Contains(sender) && mail.HasAttachments)
            	{
              		int attCounter = mail.Attachments.Count;
                //	Console.WriteLine("Attachments count: " + attCounter);
                	
                	for (int i = 0; i < attCounter; i++)
                	{
                		FileAttachment fileAttachment = mail.Attachments[i] as FileAttachment;
                		var attNameSplit = fileAttachment.Name.Split('.');
                		const string checkExt = "pdf";
                   		if (attNameSplit[1].ToLower().Contains(checkExt))
                		{
                		//  Console.WriteLine("File name is: " + attNameSplit[0]);
                		//  Console.WriteLine("Save path is: " + saveFolder);
                		//  fileAttachment.Load(saveFolder + attNameSplit[0] + ".pdf");
                		    fileAttachment.Load(save_Einv + "UNI\\NewInvoices\\" + attNameSplit[0] + ".pdf");
                		    File.Create(save_Einv + "UNI\\NewInvoices\\" + "invoice.txt").Dispose();
                		}
                	}
                	
                	                	                	                	
                	 File.Create(save_Einv + "UNI\\NewInvoices\\" + "invoice.txt").Dispose();
            //      	 var uniCount = Directory.EnumerateFiles(save_Einv + "UNI\\NewInvoices\\", "*.*", SearchOption.AllDirectories).Count();
                  	 int uniCount = new DirectoryInfo(save_Einv + "UNI\\NewInvoices\\").GetFiles().Length;
                	 Console.WriteLine("Number SE invoices processed: " + uniCount); 
		}
              	else
              	{
              		Console.WriteLine("No new mails.");
              	}
	}
        }
    }
}