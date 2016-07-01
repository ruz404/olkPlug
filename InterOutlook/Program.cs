using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;

namespace InterOutlook
{
    class Program
    {
        static void Main(string[] args)
        {
            string mailDest = "";
            string body = "";

            Console.Write("\nDestinatario: ");
            mailDest = Console.ReadLine();

            Console.Write("\nEscribir mensaje: ");
            body = Console.ReadLine();
            sendMail(mailDest, body, "SysReplay");
        }

       
        //método que responde un correo
        public static void sendMail(string mailDest, string body, string sub = "Your result", List<string> attachments = null)
        {
            Outlook.Application olkApp1 = new Outlook.Application();
            Outlook.MailItem olkMail1 = (Outlook.MailItem)olkApp1.CreateItem(Outlook.OlItemType.olMailItem);
            olkMail1.To = mailDest;
            olkMail1.CC = "";
            olkMail1.Subject = sub;
            olkMail1.Body = body;

            Console.WriteLine("¿Adjuntará algún archivo?");
            string resp = Console.ReadLine();

            if (resp != "No" && resp != "n" && resp != "no" && resp != "N" && resp != "NO")
            {
                olkMail1.Attachments.Add("c:/Users/nancyp.cruz/Downloads/Anteproyecto.pdf", Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            }


            MailMessage m = new MailMessage();
            m.From = new MailAddress("sdaf@sadf.com", "Softtron");

            if (attachments != null)
            {
                foreach (var item in attachments)
                {
                    olkMail1.Attachments.Add(item, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                }
            }
            olkMail1.Save();
            Outlook.Accounts accounts = olkApp1.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                olkMail1.Display(true);
                
            }
        }

    }


}
