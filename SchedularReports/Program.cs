using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Net.Mail;
using System.Net;
using OfficeOpenXml;
using System.Configuration;

namespace SchedularReports
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            string today = DateTime.Now.DayOfWeek.ToString();

            if(today.Trim().ToLower() == "monday")
            {
                program.TruckExecuteAndSendEmails();
                Console.WriteLine("Report Sended On "+today+": " + DateTime.Now);
            }
            
            else if (today.Trim().ToLower() == "saturday")
            {
                program.ExecuteAndSendEmails();
                Console.WriteLine("Report Sended On "+today+": " + DateTime.Now);
            }
            else
            {
                Console.WriteLine("Report Not sended because there is no day matched");
            }

            Console.ReadLine();
        }

        public async Task ExecuteAndSendEmails()
        {
            //var truckDetails = await GetTruckInOutDetails();
            var balanceCyContainer = await GetBalanceCyContainer();
            var balanceCfsCargo = await GetBalanceCfsCargo();

            // Define custom file names
            // var truckExcelFileName = $"TruckInOutDetails_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            var cyContainerExcelFileName = $"BalanceCyContainer_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            var cfsCargoExcelFileName = $"BalanceCfsCargo_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            // Generate Excel files with custom names
            //  var truckExcelFile = GenerateExcel(truckDetails, truckExcelFileName);
            var cyContainerExcelFile = GenerateExcel(balanceCyContainer, cyContainerExcelFileName);
            var cfsCargoExcelFile = GenerateExcel(balanceCfsCargo, cfsCargoExcelFileName);

            // Send emails with the named Excel file attachments
            SendEmailWithAttachments2(cyContainerExcelFile, cfsCargoExcelFile);
        }
        public async Task<SqlDataReader> GetBalanceCyContainer()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("ImportBalanceCycontainer", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            cmd.Parameters.AddWithValue("@fromDate", DBNull.Value);
            cmd.Parameters.AddWithValue("@toDate", DateTime.Now.AddDays(-1).ToShortDateString());
            cmd.Parameters.AddWithValue("@shippingAgent", "0");
            cmd.Parameters.AddWithValue("@goodsHead", "0");

            await connection.OpenAsync();
            return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
        }

        public async Task<SqlDataReader> GetBalanceCfsCargo()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("ImportBalanceCfsCargo", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            cmd.Parameters.AddWithValue("@fromDate", DBNull.Value);
            cmd.Parameters.AddWithValue("@toDate", DateTime.Now.AddDays(-1).ToShortDateString());

            await connection.OpenAsync();
            return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
        }
        public async Task TruckExecuteAndSendEmails()
        {
            var truckDetails = await GetTruckInOutDetails();
            // Define custom file names
            var truckExcelFileName = $"TruckInOutDetails_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            var truckExcelFile = GenerateExcel(truckDetails, truckExcelFileName);
            // Generate Excel files with custom names


            // Send emails with the named Excel file attachments
            SendEmailWithAttachments(truckExcelFile);
        }
        public async Task<SqlDataReader> GetTruckInOutDetails()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("usp_GetTruckInOutDetails", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            cmd.Parameters.AddWithValue("@StartDate", DateTime.Now.AddDays(-7).ToShortDateString());
            cmd.Parameters.AddWithValue("@EndDate", DateTime.Now.ToShortDateString());

            await connection.OpenAsync();
            return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
        }
        public FileInfo GenerateExcel(SqlDataReader dataReader, string fileName)
        {
            // Set the license context
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Define the folder path to save the Excel file
            string folderPath = Path.Combine(AppContext.BaseDirectory, "GeneratedReports");

            // Ensure the directory exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Generate the file path
            string filePath = Path.Combine(folderPath, fileName);

            // Create the Excel file
            FileInfo file = new FileInfo(filePath);

            // Check if the file already exists and delete it if necessary
            if (file.Exists)
            {
                file.Delete();  // Delete existing file before creating a new one
            }

            // Create the Excel package and add data
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // Add a new worksheet
                var worksheet = package.Workbook.Worksheets.Add("Report");

                // Write the header row (column names)
                for (int i = 0; i < dataReader.FieldCount; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataReader.GetName(i);
                }

                // Write the data rows
                int row = 2;  // Start from row 2 as row 1 is for headers
                while (dataReader.Read())
                {
                    for (int i = 0; i < dataReader.FieldCount; i++)
                    {
                        worksheet.Cells[row, i + 1].Value = dataReader.GetValue(i).ToString();
                    }
                    row++;
                }

                // Save the Excel file
                package.Save();
            }

            return file;  // Return the generated Excel file
        }

        private void SendEmailWithAttachments(params FileInfo[] files)
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("enquiries@aictpk.com");

                mail.To.Add("mansoor.mehmood@aictpk.com");
                mail.To.Add("javvad.qureshi@aictpk.com");
                mail.To.Add("sabir.khan@aictpk.com");
                mail.To.Add("arsalan.tariq@aictpk.com");
                mail.CC.Add("khuhronomi@gmail.com");
                mail.CC.Add("dev.team@aictpk.com");
                mail.CC.Add("saad.hassan@aictpk.com");

                //mail.To.Add("saad.hassan@aictpk.com");
                mail.CC.Add("javvad.qureshi@aictpk.com");



                mail.Subject = "Weekly Chairman Truck In/Out Report (" + DateTime.Now.AddDays(-7).ToShortDateString() + " - " + DateTime.Now.ToShortDateString() + ")";

                // Build the email body content
                string bodyContent = "<p>Please find the attached reports:</p><ul>";
                int i = 1;

                if (files.Length > 0)
                {
                    foreach (var file in files)
                    {
                        // Extract file name from the full file path
                        string fileName = file.Name.Replace("_", " ").Replace(".xlsx", "") + " Report";
                        bodyContent += $"<li>{i}. {fileName}</li>";
                        i++;
                    }
                }

                bodyContent += "</ul>";

                bodyContent += "</hr>";
                bodyContent += "<h2> Note: This is auto generated scheduled reports, no signature required.</h2>";


                mail.Body = bodyContent;
                mail.IsBodyHtml = true; // Set the body format to HTML

                foreach (var file in files)
                {
                    //if (file.Exists)
                    //{
                    mail.Attachments.Add(new Attachment(file.FullName));
                    //}
                }

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("enquiries@aictpk.com", "En@icties");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }
        }

        private void SendEmailWithAttachments2(params FileInfo[] files)
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("enquiries@aictpk.com");
                mail.To.Add("mansoor.mehmood@aictpk.com");
                mail.To.Add("javvad.qureshi@aictpk.com");
                mail.To.Add("sabir.khan@aictpk.com");
                mail.To.Add("arsalan.tariq@aictpk.com");
                mail.CC.Add("khuhronomi@gmail.com");
                mail.CC.Add("dev.team@aictpk.com");
                mail.CC.Add("saad.hassan@aictpk.com");

                //mail.To.Add("saad.hassan@aictpk.com");
                mail.CC.Add("javvad.qureshi@aictpk.com");

                mail.Subject = "Weekly Chairman Aging Report (" + DateTime.Now.AddDays(-1).ToShortDateString() + ")";

                // Build the email body content
                string bodyContent = "<p>Please find the attached reports:</p><ul>";
                int i = 1;

                if (files.Length > 0)
                {
                    foreach (var file in files)
                    {
                        // Extract file name from the full file path
                        string fileName = file.Name.Replace("_", " ").Replace(".xlsx", "") + " Report";
                        bodyContent += $"<li>{i}. {fileName}</li>";
                        i++;
                    }
                }

                bodyContent += "</ul>";

                bodyContent += "</hr>";
                bodyContent += "<h2> Note: This is auto generated scheduled reports, no signature required.</h2>";


                mail.Body = bodyContent;
                mail.IsBodyHtml = true; // Set the body format to HTML

                foreach (var file in files)
                {
                    //if (file.Exists)
                    //{
                    mail.Attachments.Add(new Attachment(file.FullName));
                    //}
                }

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("enquiries@aictpk.com", "En@icties");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }
        }

    }
}
