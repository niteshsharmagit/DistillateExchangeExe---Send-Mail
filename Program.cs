using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace DistillateExchangeExe___Send_Mail
{
    class Program
    {
        public static string FilePath = string.Empty;
        static void Main(string[] args)
        {
            try
            {
                SqlConnection con = new SqlConnection("server=3.20.32.190; database=DistillerExchange; user=sa; password=lms123; Persist Security Info=False; Connect Timeout=25000; MultipleActiveResultSets=True;");
                con.Open();

                FilePath = @"C:\DxLab - EXE\DistillateExchangeExe - Send Mail"; 
                //FilePath = @"C:\Users\Nitesh Sharma\source\repos\DistillateExchangeExe - Send Mail";

                if (!File.Exists(FilePath + "\\" + "log.txt"))
                    File.Create(FilePath + "\\" + "log.txt");

                bool isSent = false;

                string AdminIdQuery = "select Top 1 Email from tblUser Where Active = 1 and UserTypeId = 1";
                SqlCommand AdminIdcmd = new SqlCommand(AdminIdQuery, con);
                SqlDataReader AdminIdReader = AdminIdcmd.ExecuteReader();
                string AdminId = ""; //

                while (AdminIdReader.Read())
                {
                    AdminId = AdminIdReader["Email"].ToString();
                }

                SqlCommand cmd = new SqlCommand("select UserId , Email, LicenseExpiry from tblUser Where Active = 1 and UserTypeId = 2 and Status = 2", con);

                using (SqlDataReader oReader = cmd.ExecuteReader())
                {
                    if (oReader.HasRows == true)
                    {
                        Program obj = new Program();

                        while (oReader.Read())
                        {
                            DateTime LicenseExpiry = (DateTime)oReader["LicenseExpiry"];

                            SqlDataReader DateDiffReader = obj.GetDateDifference(con, LicenseExpiry);
                            while (DateDiffReader.Read())
                            {
                                if (AdminIdReader.HasRows == true)
                                {
                                    if ((int)DateDiffReader["DateDiff"] < 31 && (int)DateDiffReader["DateDiff"] > 0)
                                    {
                                            isSent = obj.SendMail("distillate.lms@gmail.com", oReader["Email"].ToString(), "odzrecwvlwkeekzu", "License Expiration Reminder", "Your license registered with DistillateEx will expire on " + LicenseExpiry + ".", AdminId);
                                    }
                                    else if ((int)DateDiffReader["DateDiff"] == 0)
                                    {
                                        isSent = obj.SendMail("distillate.lms@gmail.com", oReader["Email"].ToString(), "odzrecwvlwkeekzu", "License Expiration Reminder", "Your license registered with DistillateEx will expire today.", AdminId);
                                    }
                                    else if ((int)DateDiffReader["DateDiff"] < 0)
                                    {
                                        isSent = obj.SendMail("distillate.lms@gmail.com", oReader["Email"].ToString(), "odzrecwvlwkeekzu", "License Expiration Reminder", "Your license registered with DistillateEx has expired.", AdminId);
                                    }
                                }                                
                            }                         
                        }
                    }
                    con.Close();
                }
            }
            catch (Exception e)
            {
                using (StreamWriter w = File.AppendText(FilePath + "\\" + "log.txt"))
                    AppendLog(e.Message, w);
            }
        }
        

        public SqlDataReader GetDateDifference(SqlConnection con, DateTime LicenseExpiry)
        {
            SqlDataReader DateDiffReader = null;
            try
            {
                string DateDiffQuery = "SELECT DATEDIFF(DAY, '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + LicenseExpiry.ToString("yyyy-MM-dd HH:mm:ss") + "') AS DateDiff";
                SqlCommand DateDiffcmd = new SqlCommand(DateDiffQuery, con);
                DateDiffReader = DateDiffcmd.ExecuteReader();
            }
            catch (Exception e)
            {
                using (StreamWriter w = File.AppendText(FilePath + "\\" + "log.txt"))
                    AppendLog(e.Message, w);
            }
            return DateDiffReader;
        }

        public bool SendMail(String FromEmail, string ToEmail, string Password, string Subject, String Body, String AdminId)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress(FromEmail);
                message.To.Add(new MailAddress(ToEmail));
                message.CC.Add(new MailAddress(AdminId));
                message.Subject = Subject;

                StringBuilder body = new();
                body.AppendLine(Body);
                body.AppendLine("Please update it immediately to avoid account deactivation.");
                body.AppendLine("For any further queries contact" + " " + AdminId);
                body.AppendLine();
                body.AppendLine();
                body.AppendLine("Regards,");
                body.AppendLine("Team DistillateEx");

                message.Body = body.ToString();
                smtp.Port = 587;
                smtp.Host = "smtp.gmail.com"; //for gmail host  
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential(FromEmail, Password);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                return true;
            }
            catch (Exception ex)
            {
                using (StreamWriter w = File.AppendText(FilePath + "\\" + "log.txt"))
                    AppendLog(ex.Message, w);
                return false;
            }
        }

        private static void AppendLog(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                txtWriter.WriteLine(logMessage);
                txtWriter.WriteLine("---------------------------------------------------------------------");
            }
            catch (Exception ex)
            {
            }
        }

    }
}
