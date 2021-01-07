using System;
using System.IO;
using System.IO.Log;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Net.Mail;
using System.Net.Mime;
using Topshelf.Logging;
using CsvHelper;
using System.Data;
using System.Globalization;

namespace Paycom_File_Merge
{
    class Program
    {
        static void Main(string[] args)
        {
            //Email Notification Settings
            //List<string> recipients = new List<string>();
            //recipients.Add("ggoshert@afncorp.com");
            //recipients.Add("mblitz@afncorp.com");
            //recipients.Add("mike.nelson@afncorp.com");

            string recipients = ConfigurationManager.AppSettings["recipients"];
            string from = "ggoshert@afncorp.com";
            string subject = "Paycom File Merge Process";
            string msg = "At least 1 File is Missing!";

            //Folder & File Settings
            string sourceFolder = ConfigurationManager.AppSettings[@"sourceFolder"]; //@"C:\Paycom_Files";
            string destinationFolder = ConfigurationManager.AppSettings[@"destinationFolder"];
            string destinationFile = Path.Combine(destinationFolder, ConfigurationManager.AppSettings["outputFile"]); //@"C:\Paycom_Files\Destination\Paycom_Combined.csv";
            string archiveFolder = ConfigurationManager.AppSettings[@"archiveFolder"]; //@"C:\Paycom_Files\Archive";
            string inputFile1 = Path.Combine(sourceFolder, ConfigurationManager.AppSettings["inputFile1"]);
            string inputFile2 = Path.Combine(sourceFolder, ConfigurationManager.AppSettings["inputFile2"]);
            //string searchPattern = "*.csv";
            //char seperator = ',';


            // Specify wildcard search to match CSV files that will be combined
            //string[] filePaths = Directory.GetFiles(sourceFolder, "*.csv");
            //StreamWriter fileDest = new StreamWriter(destinationFile, true);
            try
            {
                //Make sure both files exist, otherwise send email notification
                if (File.Exists(@inputFile1) && File.Exists(@inputFile2))
                {
                    //Read first file into a CSV Object
                    var reader1 = new StreamReader(inputFile1);
                    var csv1 = new CsvReader(reader1, CultureInfo.InvariantCulture);
                    var dr1 = new CsvDataReader(csv1);

                    //Load first CSV file into a Data Table
                    var dt1 = new DataTable();
                    dt1.Load(dr1);

                    //Read second file into a CSV Object
                    var reader2 = new StreamReader(inputFile2);
                    var csv2 = new CsvReader(reader2, CultureInfo.InvariantCulture);
                    var dr2 = new CsvDataReader(csv2);

                    //Load second CSV file into a Data Table
                    var dt2 = new DataTable();
                    dt2.Load(dr2);

                    //Combine Data Tables
                    dt1.Merge(dt2);

                    //Convert Merged Data Table to CSV and write to destination
                    dt1.ToCsv(destinationFile);

                    //Archive files
                    List<string> filePaths = new List<string>();
                    filePaths.Add(inputFile1);
                    filePaths.Add(inputFile2);

                    foreach (string s in filePaths)
                    {
                        //string fileName = Path.GetFileName(s);
                        string fileName = Path.GetFileNameWithoutExtension(s);
                        string destFile = Path.Combine(archiveFolder, fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv");
                        File.Move(s, destFile);
                    }

                }

                else
                {

                    SendEmail(recipients, from, subject, msg, true);
                    //Console.WriteLine("At least 1 file is missing!");
                }

            }

            catch (Exception ex)
            {
                SendEmail(recipients, from, subject, ex.ToString(), true);

            }


            //Console.WriteLine("Hello World!");
        }

        private static string SendEmail(string to, string from, string subject, string body, bool alert)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(from);
            //for (int j = 0; j < attachments.Count; j++)
            //{
            //    mailMessage.Attachments.Add(new Attachment(attachments[j]));
            //}
            mailMessage.Subject = subject;
            mailMessage.Body = body;
            //LinkedResource LinkedImage = new LinkedResource(MRALSEmailImageFileName);
            //LinkedImage.ContentId = "MyPic";
            //LinkedImage.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);
            //AlternateView htmlView = AlternateView.CreateAlternateViewFromString("<img src=cid:MyPic>", null, "text/html");
            //htmlView.LinkedResources.Add(LinkedImage);
            //mailMessage.AlternateViews.Add(htmlView);
            SmtpClient client = new SmtpClient("mail.smtp2go.com");
            client.EnableSsl = false;
            //client.Port = 25;
            client.Port = 2525;
            //client.Credentials = new System.Net.NetworkCredential(@"afncorp\donotreply", "Alaska1234!");
            //client.Credentials = new System.Net.NetworkCredential(@"dev@afncorp.com", "t@llCanary90");
            client.Credentials = new System.Net.NetworkCredential(@"paycomfm@afncorp.com", "MjQ0MWE3NnpvbnMw");
            //for (int i = 0; i < to.Count; i++)
            //{
            //    mailMessage.To.Add(to[i]);
            //}

            mailMessage.To.Add(to);

            try
            {
                client.Send(mailMessage);
                return "mailSent";
            }
            catch (Exception ex)
            {
                //logWriter.LogWrite(ex.ToString());
                
            }
            return "NotSent";
        }


    }
}
