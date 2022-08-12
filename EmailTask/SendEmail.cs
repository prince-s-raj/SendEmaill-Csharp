using System;
using System.Net;
using System.Net.Mail;
using IronXL;



namespace EmailTask
{   
    class SendEmail
    {
        public static void Main(String[] args)
        {
            string sender;
            string password;
            string reciever ="";
            string subject = "";
            string body = "";
            Console.WriteLine("Sending Email using SMTP in C#\n");

            getSender();

            Console.WriteLine("Reciever email: type option -> |1| for Input OR |2| for file access ");
            int option = Convert.ToInt32(Console.ReadLine());
            switch (option)
            {
                case 1:
                    Console.WriteLine("Enter the Sender email address");
                    reciever = Console.ReadLine().Trim();
                    getData();
                    break;
                case 2:
                    reciever = AddEmailFromFile();
                    getData();
                    break;
            }
   
            try
            {

                SmtpClient newClient = new SmtpClient("smtp.gmail.com", 587);
                newClient.EnableSsl = true;
                newClient.UseDefaultCredentials = false;
                newClient.Credentials = new NetworkCredential(sender, password);
                newClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                MailMessage message = new MailMessage();
                message.To.Add(new MailAddress(sender));
                message.From = new MailAddress(reciever);
                message.Subject = subject;
                message.Body = body;

                newClient.Send(message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Sorry, Mail not Send : "+e.Message);
            }


            void getSender()
            {
                Console.WriteLine("Enter the Sender email address");
                sender = Console.ReadLine().Trim();
                Console.WriteLine("Enter the valid password for email ID");
                password = Console.ReadLine().Trim();
            }
            void getData() { 
                Console.WriteLine("Enter the Subject");
                subject = Console.ReadLine().Trim();
                Console.WriteLine("Enter the Email Body");
                body = Console.ReadLine().Trim();
            }
            static string AddEmailFromFile()
            {
                string data = "";
                WorkBook workbook = WorkBook.Load("C: \\Users\\iamsa\\OneDrive\\Desktop\\mails.xlsx");
                WorkSheet sheet = workbook.GetWorkSheet("sheet1");

                foreach (var cell in sheet["A1:A2"])
                {
                    data = cell.Text;
                }
                return data;
            }
        }
    }

}

