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

namespace MailSort
{
	class Program
	{
		public static void Main(string[] args)
		{

           Console.WriteLine("Connecting to Exchange");
           
     //    Production connection string and save folder(s)
           var address = ConfigurationManager.AppSettings["url"];
     //    var shared_mail = ConfigurationManager.AppSettings["shared_mail"];
           var credentials = ConfigurationManager.AppSettings["credentials"].Split(',');
           var einvoiceMail = ConfigurationManager.AppSettings["einvoiceMail"];
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
           
                                
           //Variable for switching additional output and configurations for testing
           var debug = ConfigurationManager.AppSettings["debug"];
           var test_url = ConfigurationManager.AppSettings["test_url"];
           var testMail = ConfigurationManager.AppSettings["shared_mail_test"];
           var test_credentials = ConfigurationManager.AppSettings["test_credentials"].Split(',');
           var saveFolderTesting = ConfigurationManager.AppSettings["saveFolderTesting"];
           var testSender = ConfigurationManager.AppSettings["testSender"]; //.Split(',');
           
           ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
           Mailbox mailMy = new Mailbox();
     
		          
           if (debug.Contains("false")){
           service.Credentials = new NetworkCredential(credentials[0], credentials[1], credentials[2]);
           service.Url = new Uri(address);	
           mailMy = new Mailbox(einvoiceMail);
           }
           
           if (debug.Contains("true")){
           service.Credentials = new NetworkCredential(test_credentials[0], test_credentials[1], test_credentials[2]);
           service.Url = new Uri(test_url);		
           mailMy = new Mailbox(testMail);	
           Console.WriteLine("Connected, counting unread.");
           }
     
     //      Console.WriteLine(mailMy);
           FolderId InboxID = new FolderId(WellKnownFolderName.Inbox, mailMy);
           Console.WriteLine(InboxID);
           
           Folder InboxF = Folder.Bind(service, InboxID);
                  
            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            FindItemsResults<Item> findResults = service.FindItems(InboxID, sf, new ItemView(20));
         //   Console.WriteLine("Find results are: " + findResults.TotalCount);
            
            int seCount = 0;
            int fuCount = 0;
            int itCount = 0;
            int maCount = 0;
            int otCount = 0;
            int shCount = 0;
            int utCount = 0;
            int uniCount = 0;
            
            foreach (EmailMessage mail in findResults.Items)
            {
            	mail.Load();
          //  	Console.WriteLine("Entered FOR loop.");
            	         
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		   
                		 //   Console.WriteLine("Number SE invoices processed: " + attCounter
                		}
                	}
                	
                	
                  	  seCount = new DirectoryInfo(save_Einv + "e5SE\\").GetFiles().Length;
               //   	 Console.WriteLine("Number SE invoices processed: " + seCount);
                  	 File.Create(save_Einv + "e5SE\\" + "invoice.txt").Dispose();
                  	 //var seCount = Directory.EnumerateFiles(save_Einv + "e5SE\\", "*.*", SearchOption.AllDirectories) select file).Count();
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}
                	
                	
                  	 fuCount = new DirectoryInfo(save_Einv + "e5FU\\").GetFiles().Length;
                //	 Console.WriteLine("Number FU invoices processed: " + fuCount);
                	 File.Create(save_Einv + "e5FU\\" + "invoice.txt").Dispose();
                  //	 var fuCount = Directory.EnumerateFiles(save_Einv + "e5FU\\", "*.*", SearchOption.AllDirectories).Count();
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}
                	 
                	 
                  	 itCount = new DirectoryInfo(save_Einv + "e5IT\\").GetFiles().Length;
               // 	 Console.WriteLine("Number IT invoices processed: " + itCount);
                	 File.Create(save_Einv + "e5IT\\" + "invoice.txt").Dispose();
              //    	 var itCount = Directory.EnumerateFiles(save_Einv + "e5IT\\", "*.*", SearchOption.AllDirectories).Count();
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}
                	
                	 
                  	 maCount = new DirectoryInfo(save_Einv + "e5MA\\").GetFiles().Length;
                //	 Console.WriteLine("Number MA invoices processed: " + maCount); 
                	 File.Create(save_Einv + "e5MA\\" + "invoice.txt").Dispose();
                 // 	 var maCount = Directory.EnumerateFiles(save_Einv + "e5MA\\", "*.*", SearchOption.AllDirectories).Count();
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}
                	
                	 
                  	 otCount = new DirectoryInfo(save_Einv + "e5OT\\").GetFiles().Length;
               // 	 Console.WriteLine("Number OT invoices processed: " + otCount);               	
                	 File.Create(save_Einv + "e5OT\\" + "invoice.txt").Dispose();
                  //	 var otCount = Directory.EnumerateFiles(save_Einv + "e5OT\\", "*.*", SearchOption.AllDirectories).Count();
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}

                	 
                  	 shCount = new DirectoryInfo(save_Einv + "e5SH\\").GetFiles().Length;
                //	 Console.WriteLine("Number SH invoices processed: " + shCount);                	
                	 File.Create(save_Einv + "e5SH\\" + "invoice.txt").Dispose();
                 // 	 var shCount = Directory.EnumerateFiles(save_Einv + "e5SH\\", "*.*", SearchOption.AllDirectories).Count();               	
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                   		}
                	}
                	
                	
                  	 utCount = new DirectoryInfo(save_Einv + "e5UT\\").GetFiles().Length;
               // 	 Console.WriteLine("Number UT invoices processed: " + utCount);                 	                	
                	 File.Create(save_Einv + "e5UT\\" + "invoice.txt").Dispose();
          //        	 var utCount = Directory.EnumerateFiles(save_Einv + "e5UT\\", "*.*", SearchOption.AllDirectories).Count(); 
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
                		    mail.IsRead = true;
                		    mail.Update(ConflictResolutionMode.AlwaysOverwrite);
                		}
                	}
                	
                	 
                  	 uniCount = new DirectoryInfo(save_Einv + "UNI\\NewInvoices\\").GetFiles().Length;
              //  	 Console.WriteLine("Number UNI invoices processed: " + uniCount);                	                	                	
                	 File.Create(save_Einv + "UNI\\NewInvoices\\" + "invoice.txt").Dispose();
            //      	 var uniCount = Directory.EnumerateFiles(save_Einv + "UNI\\NewInvoices\\", "*.*", SearchOption.AllDirectories).Count();
		}
              	if (seCount > 1){
              		
              		Console.WriteLine("SE invoices processed.");
              	}
              	
              	if (fuCount > 1){
              		
              		Console.WriteLine("FU invoices processed.");
              	}
              	
              	if (itCount > 1){
              		
              		Console.WriteLine("IT invoices processed.");
              	}
              	
              	if (maCount > 1){
              		
              		Console.WriteLine("MA invoices processed.");
              	}
              	
              	if (otCount > 1){
              		
              		Console.WriteLine("MA invoices processed.");
              		
              	}
              	
              	if (shCount > 1){
              		
              		Console.WriteLine("SH invoices processed.");
              	}
              	
              	if (utCount > 1){
              		
              		Console.WriteLine("UT invoices processed.");
              	}
              	
              	if (uniCount > 1){
              		
              		Console.WriteLine("UNI invoices processed.");
              	}
              	
              	             	
/*              	else
              	{
              		Console.WriteLine("No new mails.");
              	}*/
	}
                Console.WriteLine("");
            
              	if (seCount > 1){
              		
              		Console.WriteLine("SE invoices processed.");
              	}
              	
              	if (fuCount > 1){
              		
              		Console.WriteLine("FU invoices processed.");
              	}
              	
              	if (itCount > 1){
              		
              		Console.WriteLine("IT invoices processed.");
              	}
              	
              	if (maCount > 1){
              		
              		Console.WriteLine("MA invoices processed.");
              	}
              	
              	if (otCount > 1){
              		
              		Console.WriteLine("MA invoices processed.");
              		
              	}
              	
              	if (shCount > 1){
              		
              		Console.WriteLine("SH invoices processed.");
              	}
              	
              	if (utCount > 1){
              		
              		Console.WriteLine("UT invoices processed.");
              	}
              	
              	if (uniCount > 1){
            	              		
              		Console.WriteLine("UNI invoices processed.");
              	}
            else {
            	
            	Console.WriteLine("No new invoices.");
            }
              Console.ReadLine();
        }
    }
}