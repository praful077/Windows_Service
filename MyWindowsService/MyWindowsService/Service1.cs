using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MyWindowsService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); // name space(using System.Timers;)
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000; //number in milisecinds
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service is recall at " + DateTime.Now);

            //string documentPath = @"C:/my/demo/MyWindowsService/MyWindowsService/GeneratedDocument.docx";
            // GenerateDocument(documentPath, "This is the generated document.");
            // Generate the Word document with simple data
            byte[] documentBytes = GenerateDocument("This is the generated document. and send by Praful");

            // Send the email with the generated document as an attachment
            string recipient = "praful.parmar@internal.mail";
            string subject = "Attendance is penting";
            string body = "This is a test email.";
            //string filePath = documentPath;

            SendEmail(recipient, subject, body, documentBytes);
            WriteToFile("Service sent an email at 1 " + DateTime.Now);
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                //InstallUtil.exe C:\my\demo\MyWindowsService\MyWindowsService\bin\Debug\MyWindowsService.exe
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private byte[] GenerateDocument(string documentContent)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph paragraph = body.AppendChild(new Paragraph());
                    Run run = paragraph.AppendChild(new Run());
                    run.AppendChild(new Text(documentContent));

                    wordDocument.MainDocumentPart.Document.Save();
                }
                WriteToFile("Service a Generate document  at " + DateTime.Now);

                return memoryStream.ToArray();
            }
        }

        //private void GenerateDocument(string documentPath, string documentContent)
        //{
        //    // Create a new Word document
        //    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(documentPath, WordprocessingDocumentType.Document))
        //    {
        //        // Add a new main document part
        //        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

        //        // Create the document structure
        //        mainPart.Document = new Document();
        //        Body body = mainPart.Document.AppendChild(new Body());
        //        Paragraph paragraph = body.AppendChild(new Paragraph());
        //        Run run = paragraph.AppendChild(new Run());
        //        run.AppendChild(new Text(documentContent));

        //        // Save the document
        //        wordDocument.MainDocumentPart.Document.Save();
        //        WriteToFile("Service a Generate document  at " + DateTime.Now);
        //    }
        //}

        private void SendEmail(string recipient, string subject, string body, byte[] documentBytes)
        {
           
            WriteToFile("Service sent an email at 2 " + DateTime.Now);

            var emailfrom = new MailAddress("prafulparmar1178@gmail.com");
                var toEmail = new MailAddress(recipient);
                var frompwd = "fpbdasefiupbffsx";

            string templatePath = @"C:/my/demo/MyWindowsService/MyWindowsService/Gmail_Template/Birthday_Wish.html";
            string emailBody;

            using (StreamReader reader = new StreamReader(templatePath))
            {
                emailBody = reader.ReadToEnd();
            }

            Attachment attachment = null;

            if (documentBytes != null && documentBytes.Length > 0)
            {
                MemoryStream documentStream = new MemoryStream(documentBytes);
                attachment = new Attachment(documentStream, "GeneratedDocument.docx");
            }

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(emailfrom.Address, frompwd)
            };

            WriteToFile("Service sent an email at 3 " + DateTime.Now);
            MailMessage message = new MailMessage(emailfrom, toEmail)
            {
                Subject = subject,
                Body = emailBody + "\n\n" + body,
                IsBodyHtml = true
            };

            if (attachment != null)
            {
                message.Attachments.Add(attachment);
            }

            smtp.Send(message);

            WriteToFile("Service sent an email at 4 " + DateTime.Now);



        }

    }
}
