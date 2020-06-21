using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace EmailReader
{
    class Program
    {
        private static IConfigurationRoot Configuration { get; set; }
        public static void Main(string[] args)
        {
            Configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false).Build();

            var host = Configuration["iCloudEmailConfig:Host"];
            int port = Convert.ToInt32(Configuration["iCloudEmailConfig:Port"]);
            var userName = Configuration["iCloudEmailConfig:UserName"];
            var password = Configuration["iCloudEmailConfig:Password"];
            

            List<BrokerageNote> brokerageNotes = new List<BrokerageNote>() { };

            using (var client = new ImapClient())
            {
                client.Connect(host, port, true);

                client.Authenticate(userName, password);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadOnly);

                DateTime today = DateTime.Now;

                var clear = inbox.Search(SearchQuery.FromContains("noreply@clear.com.br").And(SearchQuery.DeliveredAfter(today.AddDays(-30))));

                Console.WriteLine("Total messages: {0}", clear.Count);

                for (int index = 0; index < clear.Count; index++)
                {
                    var message = inbox.GetMessage(clear[index]);

                    var messageDate = message.Date.DateTime.Date;

                    string pdfLink = GetPdfLinks(message.HtmlBody);

                    string fileName = $"nota-de-corretagem-{messageDate.Day - 1}-{messageDate.Month}-{messageDate.Year}.pdf";

                    brokerageNotes.Add(new BrokerageNote { Date = messageDate, PdfLink = pdfLink, FileName = fileName });

                    Console.WriteLine($"pdfLink: {pdfLink}");
                    Console.WriteLine($"messageDate: {messageDate}");
                }

                client.Disconnect(true);
            }

            DownloadPdfFiles(brokerageNotes);

        }

        private static void DownloadPdfFiles(List<BrokerageNote> brokerageNotes)
        {
            string pdfPassword = Configuration["PdfPassword"];
            foreach (var bN in brokerageNotes)
            {
                using (WebClient client = new WebClient())
                {
                    var pdfBytes = client.DownloadData(bN.PdfLink);

                    PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfBytes, pdfPassword);

                    loadedDocument.Security.Permissions = PdfPermissionsFlags.Default;
                    loadedDocument.Security.UserPassword = string.Empty;
                    loadedDocument.Security.OwnerPassword = string.Empty;

                    var pdfWithoutPassword = File.Create(@$"C:\Users\Estanislau\Desktop\notas\{bN.FileName}");
                    loadedDocument.Save(pdfWithoutPassword);
                    loadedDocument.Close(true);
                    client.Dispose();

                    Console.WriteLine($"{bN.FileName} pdfFile downloaded");
                }
            }
        }

        private static string GetPdfLinks(string htmlBody)
        {
            var html = new HtmlDocument();

            html.LoadHtml(htmlBody);

            var link = html.DocumentNode.SelectNodes("//a[@href]").FirstOrDefault(l => l.InnerText.Contains("Nota de Negocia"));

            if(link != null)
            {
                string hrefValue = link.GetAttributeValue("href", string.Empty);
                return hrefValue;
            }

            return string.Empty;
        }

    }
}
