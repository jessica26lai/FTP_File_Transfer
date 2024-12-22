using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.IO;
using System.Data;
using System.Web;
using System.Data.OleDb;
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Data.SqlClient;
using System.Security;

namespace FTP_File_Transfer
{
    class Program
    {
        private bool emailSent = false;

        static void Main(string[] args)
        {
            Program p = new Program();

            String src_method = "local";
            String dest_method = "ftp";

            try
            {
                Console.WriteLine("Scheduler starts");

                if (args.Length != 0)
                {
                    src_method = args[0];
                    dest_method = args[1];
                }

                if (src_method == "local")
                {

                    DirectoryInfo directorySource = new DirectoryInfo(Properties.Settings.Default.SourcePath);

                    if (directorySource.GetFiles().Length > 0)
                    {

                        if (dest_method == "ftp")
                        {

                            using (var sftp = new SftpClient(Properties.Settings.Default.DestHostName, Properties.Settings.Default.DestPort, Properties.Settings.Default.DestUserName, Properties.Settings.Default.DestPassword))
                            {
                                sftp.Connect();

                                var files = Directory.GetFiles(Properties.Settings.Default.SourcePath);
                                foreach (var file in files)
                                {
                                    using (Stream file1 = new FileStream(file, FileMode.Open))
                                    {
                                        string filename = Path.GetFileName(file);

                                        string destPath = Path.Combine(Properties.Settings.Default.DestPathFTP, filename);
                                        destPath = destPath.Replace("\\", "/");
                                        sftp.UploadFile(file1, destPath, false, null);

                                        Console.WriteLine("Upload of {0} succeeded", filename);
                                        p.LogTable.Rows.Add("UPLOAD_FILES_TO_FTP", "Upload of " + filename + " succeeded", true);

                                        

                                    }

                                    p.backupFile(file);

                                }

                                sftp.Disconnect();

                            }

                        }
                        else if (dest_method == "local")
                        {

                            var files = Directory.GetFiles(Properties.Settings.Default.SourcePath);
                            foreach (var file in files)
                            {
                                if (!Directory.Exists(Properties.Settings.Default.DestPath))
                                {
                                    Directory.CreateDirectory(Properties.Settings.Default.DestPath);
                                }

                                string filename = Path.GetFileName(file);
                                string destFilePath = Path.Combine(Properties.Settings.Default.DestPath, filename);

                                File.Copy(file, destFilePath, false);

                                Console.WriteLine("Folder Move of {0} succeeded", filename);
                                p.LogTable.Rows.Add("MOVE_FILES_TO_FOLDER", "Move of " + filename + " succeeded", true);

                                p.backupFile(file);
                            }

                        }
                        else
                        {
                            throw new Exception("Invalid destination parameter.");
                        }


                    }
                    else
                    {
                        p.LogTable.Rows.Add("GET_LOCAL_FILES", "No files found", true);
                    }
                }
                else if (src_method == "ftp")
                {

                    if (dest_method == "ftp")
                    {

                        using (var sftp = new SftpClient(Properties.Settings.Default.SourceHostName, Properties.Settings.Default.SourcePort, Properties.Settings.Default.SourceUserName, Properties.Settings.Default.SourcePassword))
                        {
                            sftp.Connect();

                            string SourcePathFTP = Properties.Settings.Default.SourcePathFTP;

                            var files = sftp.ListDirectory(SourcePathFTP, null);
                            foreach (var file in files)
                            {
                                string remoteFileName = file.Name;

                                string destFilePath = Path.Combine(Properties.Settings.Default.FTPTransferTempPath, remoteFileName);

                                using (Stream file1 = File.OpenWrite(destFilePath))
                                {
                                    sftp.DownloadFile(SourcePathFTP +'/'+ remoteFileName, file1, null);

                                    Console.WriteLine("Download of {0} succeeded", remoteFileName);
                                    p.LogTable.Rows.Add("DOWNLOAD_FILES_FROM_FTP", "Download of " + remoteFileName + " succeeded", true);

                                    sftp.DeleteFile(SourcePathFTP + '/' + remoteFileName);
                                }
                                
                            }

                            sftp.Disconnect();

                        }

                        using (var sftp = new SftpClient(Properties.Settings.Default.DestHostName, Properties.Settings.Default.DestPort, Properties.Settings.Default.DestUserName, Properties.Settings.Default.DestPassword))
                        {
                            sftp.Connect();

                            var files = Directory.GetFiles(Properties.Settings.Default.FTPTransferTempPath);
                            foreach (var file in files)
                            {
                                using (Stream file1 = new FileStream(file, FileMode.Open))
                                {
                                    string filename = Path.GetFileName(file);

                                    string destPath = Path.Combine(Properties.Settings.Default.DestPathFTP, filename);
                                    destPath = destPath.Replace("\\", "/");
                                    sftp.UploadFile(file1, destPath, false, null);

                                    Console.WriteLine("Upload of {0} succeeded", filename);
                                    p.LogTable.Rows.Add("UPLOAD_FILES_TO_FTP", "Upload of " + filename + " succeeded", true);
                                }

                                p.backupFile(file);

                            }

                            sftp.Disconnect();

                        }


                    }
                    else if (dest_method == "local")
                    {

                        using (var sftp = new SftpClient(Properties.Settings.Default.SourceHostName, Properties.Settings.Default.SourcePort, Properties.Settings.Default.SourceUserName, Properties.Settings.Default.SourcePassword))
                        {
                            sftp.Connect();

                            string SourcePathFTP = Properties.Settings.Default.SourcePathFTP;

                            var files = sftp.ListDirectory(SourcePathFTP, null);
                            foreach (var file in files)
                            {
                                string remoteFileName = file.Name;

                                string destFilePath = Path.Combine(Properties.Settings.Default.DestPath, remoteFileName);

                                if (File.Exists(destFilePath))
                                {
                                    throw new Exception("Local file exists.");
                                }
                                else
                                {

                                    using (Stream file1 = File.OpenWrite(destFilePath))
                                    {
                                        sftp.DownloadFile(SourcePathFTP + '/' + remoteFileName, file1, null);

                                        Console.WriteLine("Download of {0} succeeded", remoteFileName);
                                        p.LogTable.Rows.Add("DOWNLOAD_FILES_FROM_FTP", "Download of " + remoteFileName + " succeeded", true);

                                        sftp.DeleteFile(SourcePathFTP + '/' + remoteFileName);
                                    }
                                }
                            }

                            sftp.Disconnect();

                        }

                    }
                    else
                    {
                        throw new Exception("Invalid destination parameter.");
                    }

                }
                else
                {
                    throw new Exception("Invalid source parameter.");
                }

                Console.WriteLine("Program completed");
                p.LogTable.Rows.Add("MAIN_THREAD", "Program completed", true);

            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);

                p.LogTable.Rows.Add("MAIN_THREAD", "ERROR: " + e.ToString(), false);
            }
            finally
            {
                p.SaveLog();

                if (p.emailSent)
                {
                    using (FileStream fs = File.Create(Properties.Settings.Default.installationPath + "\\EmailSent"))
                    {
                        byte[] info = new UTF8Encoding(true).GetBytes("");
                        fs.Write(info, 0, info.Length);
                    }

                }
                Console.WriteLine("Scheduler ended");

                //Console.ReadKey();
            }
        }

        public DataTable _logTable;
        public DataTable LogTable
        {
            get
            {
                if (_logTable == null)
                {
                    _logTable = new DataTable("Log");
                    _logTable.Columns.Add("SystemName", typeof(string));
                    _logTable.Columns.Add("Message", typeof(string));
                    _logTable.Columns.Add("Success", typeof(bool));
                }

                return _logTable;
            }
        }

        public int Count_Okay
        {
            get { return LogTable.AsEnumerable().Count(x => x.Field<bool>("Success")); }
        }

        public int Count_Failed
        {
            get { return LogTable.AsEnumerable().Count(x => !x.Field<bool>("Success")); }
        }

        private void SendEmail(string attachmentPath, int mailType)
        {
            if (!checkEmailIsSent())
            {

                //mailType: 0- scheduler errors, 1 - Success

                MailMessage mMailMessage = new MailMessage();

                try
                {
                    string from = Properties.Settings.Default.EmailFrom;


                    if (!string.IsNullOrEmpty(from))
                    {
                        mMailMessage.From = new MailAddress(from);

                        //StringBuilder sb = new StringBuilder();
                        //foreach (DataRow dr in LogTable.Rows)
                        //{
                        //    sb.AppendFormat("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>",
                        //        SecurityElement.Escape(dr[0].ToString()),
                        //        SecurityElement.Escape(dr[1].ToString()),
                        //        SecurityElement.Escape(dr[2].ToString()));
                        //}

                        if (mailType == 1)
                        {
                            mMailMessage.Subject = Properties.Settings.Default.EmailSubjectSuccess;
                            mMailMessage.Subject = "[ALERT]" + mMailMessage.Subject;
                            mMailMessage.Body = string.Format(HttpUtility.HtmlDecode(Properties.Settings.Default.EmailBodySuccess),
                              DateTime.Now,
                              TimeSpan.FromMilliseconds(double.Parse(Properties.Settings.Default.SourceTimeLimit)).TotalMinutes);

                            string toEmail = Properties.Settings.Default.SupportEmailTo;

                            foreach (var item in toEmail.Split(new char[] { ';' }))
                            {
                                if (item.Length != 0)
                                    mMailMessage.To.Add(item);
                            }

                        }
                        else
                        {
                            mMailMessage.Subject = Properties.Settings.Default.EmailSubjectError;
                            mMailMessage.Subject = "[FAILED]" + mMailMessage.Subject;
                            //mMailMessage.Body = string.Format(HttpUtility.HtmlDecode(Properties.Settings.Default.EmailBodySchedulerError), DateTime.Now, sb);
                            mMailMessage.Body = string.Format(HttpUtility.HtmlDecode(Properties.Settings.Default.EmailBodySchedulerError),DateTime.Now);

                            string toEmail = Properties.Settings.Default.SupportEmailTo;

                            foreach (var item in toEmail.Split(new char[] { ';' }))
                            {
                                if (item.Length != 0)
                                    mMailMessage.To.Add(item);
                            }
                        }

                        //mMailMessage.Body = "<style>table, th, td {border: 1px solid black;padding:5px 5px 5px 5px}</style>" + mMailMessage.Body;
                        //<table><th>Function</th><th>Message</th><th>Success</th>{1}</table><br/>
                        mMailMessage.Body = mMailMessage.Body;

                        mMailMessage.IsBodyHtml = true;

                        if (!string.IsNullOrEmpty(attachmentPath))
                        {

                            if (mailType > 0)
                            {
                                DirectoryInfo di = new DirectoryInfo(attachmentPath);

                                foreach (FileInfo fileInfo in di.GetFiles())
                                {

                                    mMailMessage.Attachments.Add(new Attachment(fileInfo.FullName));

                                }
                            }
                            else
                            {
                                mMailMessage.Attachments.Add(new Attachment(attachmentPath));
                            }
                        }

                        mMailMessage.Priority = MailPriority.Normal;

                        SmtpClient mSmtpClient = new SmtpClient(Properties.Settings.Default.SMTP_Server);
                        mSmtpClient.Send(mMailMessage);

                        LogTable.Rows.Add("SEND_EMAIL", string.Format("Time: {0} - Done", DateTime.Now.ToString()), true);

                        if (mailType == 0)
                        {
                            emailSent = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error:{0}", ex.ToString());
                    LogTable.Rows.Add("SEND_EMAIL", "ERROR: " + ex.ToString(), false);
                }
            }

        }

        public void SaveLog()
        {
            try
            {
                string mPath = Path.Combine(Properties.Settings.Default.installationPath, "Log\\");
                if (!Directory.Exists(mPath))
                {
                    Directory.CreateDirectory(mPath);
                }

                string filePath = string.Format(mPath + "{0:yyyyMMdd}.log", DateTime.Today);

                string firstErrorMsg = "";
                Boolean errorMarked = false;

                StringBuilder sb = new StringBuilder("");
                foreach (DataRow dr in LogTable.Rows)
                {
                    sb.Append(string.Format("[{0}] {1}", dr["SystemName"], dr["Message"]) + "\r\n");

                    if (!errorMarked &&(!((bool)dr["Success"]))){
                        firstErrorMsg = (String)dr["Message"];
                        errorMarked = true;
                    }
                }
                sb.Append("===========================\r\n");
                sb.Append(string.Format("-- End Log {0:HH:mm:ss}\r\n", DateTime.Now));
                sb.Append("===========================\r\n");


                if (!File.Exists(filePath))
                {
                    using (StreamWriter sw = File.CreateText(filePath))
                    {
                        sw.Write(sb.ToString());
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filePath))
                    {
                        sw.Write(sb.ToString());
                    }
                }
                if (Count_Failed > 0)
                {
                    //SendEmail(filePath, 0);

                    //insertLogIntoDb(firstErrorMsg, false);

                }
                else
                {
                    if (Properties.Settings.Default.sendSuccessEmail == "Y")
                    {
                        //SendEmail(null, 1);
                    }

                    //insertLogIntoDb("Scheduler Run Completed", true);

                }

            }
            catch(Exception e)
            {
                Console.WriteLine("Error:{0}", e.ToString());
                //if (Count_Failed > 0)
                    //SendEmail(e.Message, 0);
            }
        }

        public bool checkEmailIsSent()
        {
            String filepath = Properties.Settings.Default.installationPath + "\\EmailSent";

            if (File.Exists(filepath))
            {
                FileInfo fileInfo = new FileInfo(filepath);

                double timeDiff = (DateTime.Now - fileInfo.LastWriteTime).TotalHours;
                if (timeDiff > double.Parse(Properties.Settings.Default.EmailSendingTimeLimit))
                {
                    fileInfo.Delete();
                    return false;
                }
                else
                {
                    return true;
                }

            }
            return false;
        }

        public bool backupFile(string srcFilePath)
        {
            try
            {
                if (!Directory.Exists(Properties.Settings.Default.BackupPath))
                {
                    Directory.CreateDirectory(Properties.Settings.Default.BackupPath);
                }

                string filename = Path.GetFileName(srcFilePath);
                string destFilePath = Path.Combine(Properties.Settings.Default.BackupPath, filename);

                File.Move(srcFilePath, destFilePath);
                Console.WriteLine("{0} was moved to backup folder.", filename);

                return true;

            }catch(Exception ex){

                Console.WriteLine("Error:{0}", ex.ToString());
                LogTable.Rows.Add("BACKUP_FILE", "ERROR: " + ex.ToString(), false);
                return false;

            }

        }

        public bool insertLogIntoDb(string message, bool success)
        {

            try
            {
                using (SqlConnection conn = new SqlConnection(Properties.Settings.Default.ConnectionString))
                {
                    conn.Open();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "INSERT INTO [dbo].[FTP_FileTransferLog]([SystemName],[Message],[Success],[LogTime])" +
                                        "VALUES(@SystemName,@Message,@Success,@LogTime)";

                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.AddWithValue("@SystemName", Properties.Settings.Default.ApplicationName);
                    cmd.Parameters.AddWithValue("@Message", message);
                    cmd.Parameters.AddWithValue("@Success", success ? 1 : 0);
                    cmd.Parameters.AddWithValue("@LogTime", DateTime.Now);


                    cmd.ExecuteNonQuery();

                    LogTable.Rows.Add("INSERT_LOG_INTO_DB", "Done", true);

                    conn.Close();

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:{0}", ex.ToString());
                LogTable.Rows.Add("INSERT_LOG_INTO_DB", "ERROR: " + ex.ToString(), false);
                return false;
            }
        }


    }
}
