using System;
using System.IO;
using System.Data;
using Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace CRAutomatization
{
    internal class Program
    {
        static Application outlook;
        static void Main(string[] args)
        {
            outlook = new Application();
            NameSpace ns = outlook.GetNamespace("MAPI");
            MAPIFolder inbox = ns.GetSharedDefaultFolder(ns.CreateRecipient("bit-bce-salesteam@bce.bitclub.hu") ,OlDefaultFolders.olFolderInbox);
            Items items = inbox.Items.Restrict("[ReceivedTime] >= '01/01/2020'");

            DataTable dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Sender", typeof(string));
            dt.Columns.Add("To", typeof(string));
            dt.Columns.Add("Subject", typeof(string));
            dt.Columns.Add("Partner", typeof(string));

            foreach (MailItem email in items) {
                DataRow dr = dt.NewRow();
                dr["Date"] = email.SentOn;
                dr["Sender"] = GetEmailAddress(email.Sender);

                // extract email addresses for "To" recipients
                string toEmails = "";
                foreach (Recipient recipient in email.Recipients)
                {
                    if (recipient.Type == (int)OlMailRecipientType.olTo)
                    {
                        toEmails += GetEmailAddress(recipient.AddressEntry) + "; ";
                    }
                }
                dr["To"] = toEmails.TrimEnd(';');

                dr["Subject"] = email.Subject;

                //format to uppercase and keep domain for partners
                if (dr["Sender"].ToString().EndsWith("bce.bitclub.hu"))
                {
                    dr["To"] = GetFirstNonBIT(dr["To"].ToString());
                    dr["Sender"] = dr["Sender"].ToString().Split('@')[0].Replace(".", " ").ToUpper();
                    dr["To"] = GetDomain(dr["To"].ToString());
                    dr["Partner"] = dr["To"];

                } else {
                    dr["To"] = GetFirstBIT(dr["To"].ToString());
                    dr["To"] = dr["To"].ToString().Split('@')[0].Replace(".", " ").ToUpper();
                    dr["Sender"] = GetDomain(dr["Sender"].ToString());
                    dr["Partner"] = dr["Sender"];
                }

                // Check if all necessary columns are filled before adding the row
                if (!string.IsNullOrEmpty(dr["Sender"].ToString()) &&
                    !string.IsNullOrEmpty(dr["To"].ToString()) &&
                    !string.IsNullOrEmpty(dr["Subject"].ToString()) &&
                    !string.IsNullOrEmpty(dr["Date"].ToString()) &&
                    !string.IsNullOrEmpty(dr["Partner"].ToString()))
                {
                    dt.Rows.Add(dr);
                }

            }

            WriteToExcel(dt, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),"email_history.xlsx"));
        }

        private static string GetDomain(string email) {
            if (string.IsNullOrEmpty(email))
                return email;

            if (email.EndsWith("@gmail.com"))
            {
                email = email.Split('@')[0];
                email.Replace(".-_", " ").ToLower();
                return email;
            }
            else {
                email = email.Split('@')[1];
                email = email.Replace(".", " ");
                email = email.Substring(0, email.LastIndexOf(' ')).ToLower();
                return email;
            }
        }

        private static string GetFirstNonBIT(string recipients)
        {
            string[] recipientList = recipients.Split(';');
            string nonBceEmail = "";

            foreach (string recipient in recipientList)
            {
                if (recipient.Trim().EndsWith("@bce.bitclub.hu"))
                {
                    continue;
                }

                nonBceEmail = recipient.Trim();
                break;
            }

            return nonBceEmail;
        }

        private static string GetFirstBIT(string recipients)
        {
            string[] recipientList = recipients.Split(';');

            foreach (string recipient in recipientList)
            {
                if (recipient.Trim().EndsWith("@bce.bitclub.hu"))
                {
                    return recipient.Trim();
                }

            }

            return "";
        }

        static string GetEmailAddress(AddressEntry recipient)
        {
            try
            {
                // Check if the recipient is an Exchange account
                if (recipient.Type == "EX")
                {
                    return recipient.GetExchangeUser().PrimarySmtpAddress;
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            return recipient.Address;
        }

        static void WriteToExcel(DataTable dt, string filePath)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            for (int i = 0; i < dt.Columns.Count; i++) {
                worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }

            for (int i = 0; i < dt.Rows.Count; i++) {
                for (int j = 0; j < dt.Columns.Count; j++) {
                    worksheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    
                }

            }

            workbook.SaveAs(filePath);
            workbook.Close();
            excel.Quit();
        }
    }
}
