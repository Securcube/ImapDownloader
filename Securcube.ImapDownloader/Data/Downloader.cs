using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using MimeKit.IO;
using SecurCube.ImapDownloader.Data;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Securcube.ImapDownloader.Data
{
    class Downloader
    {

        internal static async Task<List<EmailFolder>> GetAllFoldersAsync(DataContext dc)
        {
            return await Task.Run(() => GetAllFolders(dc));
        }

        internal static List<EmailFolder> GetAllFolders(DataContext dc)
        {

            // The default port for IMAP over SSL is 993.
            using (ImapClient client = new ImapClient())
            {

                client.ServerCertificateValidationCallback = (s, c, h, ee) => true;

                client.Connect(dc.HostName, dc.Port, dc.UseSSL);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(dc.UserName, dc.UserPassword);

                dc.DestinationFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\ExportImap\\" + dc.UserName + "\\";

                var myFodlers = new List<EmailFolder>();

                var folders = client.GetFolders(client.PersonalNamespaces[0], StatusItems.Count | StatusItems.UidNext | StatusItems.UidValidity | StatusItems.Unread);

                foreach (var folder in folders)
                {

                    myFodlers.Add(new EmailFolder()
                    {
                        Folder = folder.FullName,
                        Messages = folder.Count,
                        NextUID = folder.UidNext != null ? folder.UidNext.Value.Id : 0,
                        Selected = !folder.Attributes.HasFlag(FolderAttributes.All) && !folder.Attributes.HasFlag(FolderAttributes.NonExistent),
                        Flags = folder.Attributes,
                        Unread = folder.Unread,
                    });

                }

                dc.EmailFolders = myFodlers;

                return myFodlers;

            }

        }

        // Check connection download speed: first param = 
        private static ConcurrentBag<Tuple<DateTime, double, long>> DownloadSpeed = new ConcurrentBag<Tuple<DateTime, double, long>>();

        /// <summary>
        /// Returns the download speed from the beginning and on the latest 30 sec.
        /// </summary>
        /// <returns></returns>
        internal static Tuple<double, double> getDownloadSpeed()
        {

            double s1 = 0, s2 = 0;

            // Check speed last 30 seconds
            var dt = DateTime.Now.AddSeconds(-30);
            var items = DownloadSpeed.Where(o => o.Item1 > dt);
            var downloadedData = items.Sum(o => o.Item3);
            var downloadedTime = items.Sum(o => o.Item2);
            if (downloadedTime != 0)
            {
                s1 = downloadedData / downloadedTime * 1000;
            }

            // check spped from beginning
            downloadedData = DownloadSpeed.Sum(o => o.Item3);
            downloadedTime = DownloadSpeed.Sum(o => o.Item2);
            if (downloadedTime != 0)
            {
                s2 = downloadedData / downloadedTime * 1000;
            }

            return new Tuple<double, double>(s1, s2);

        }


        internal static int TotalMails = 0;
        internal static int ProgressMails = 0;

        internal async static Task<long> DownloadMailsAsync(DataContext dc)
        {
            return await Task.Run(() => DownloadMails(dc));
        }

        private static long DownloadMails(DataContext dc)
        {

            List<string> log = new List<string>();
            List<string> logError = new List<string>() { "", "" };

            DateTime dtStart = DateTime.UtcNow;

            log.Add("");
            log.Add("LOGIN PARAMS");
            log.Add("Host name: " + dc.HostName + " (resolved ip = " + System.Net.Dns.GetHostEntry(dc.HostName).AddressList.FirstOrDefault() + ") ");
            log.Add("Port: " + dc.Port);
            log.Add("UseSSL: " + dc.UseSSL.ToString());
            log.Add("User name: " + dc.UserName);
            log.Add("User password: " + dc.UserPassword);
            log.Add("");

            long totalMessagesDownloaded = 0;

            // clone the list
            TotalMails = dc.EmailFolders.Where(o => o.Selected).Sum(o => o.Messages);

            Parallel.ForEach(dc.EmailFolders, new ParallelOptions() { MaxDegreeOfParallelism = dc.ConcurrentThreads }, (folder) =>
            {

                if (folder.Selected == false)
                    return;

                // The default port for IMAP over SSL is 993.
                using (ImapClient client = new ImapClient())
                {

                    client.ServerCertificateValidationCallback = (s, c, h, ee) => true;

                    try
                    {
                        client.Connect(dc.HostName, dc.Port, dc.UseSSL);
                    }
                    catch (ImapProtocolException)
                    {
                        // try twice
                        System.Threading.Thread.Sleep(100);
                        client.Connect(dc.HostName, dc.Port, dc.UseSSL);
                    }

                    // Note: since we don't have an OAuth2 token, disable
                    // the XOAUTH2 authentication mechanism.
                    client.AuthenticationMechanisms.Remove("XOAUTH2");

                    client.Authenticate(dc.UserName, dc.UserPassword);

                    ImapFolder imapFodler = (ImapFolder)client.GetFolder(folder.Folder);

                    folder.IsDownloading = true;

                    folder.DownloadedItems = 0;

                    string destFolder = Path.Combine(dc.DestinationFolder, folder.Folder.Replace(imapFodler.DirectorySeparator, '\\'));

                    // If the folder already exsist I have do delete it
                    if (!dc.MergeFolders)
                        if (Directory.Exists(destFolder))
                            Directory.Delete(destFolder, true);

                    if (!Directory.Exists(destFolder))
                        Directory.CreateDirectory(destFolder);

                    string messageIdSafeName = "";

                    int downloadedEmails = 0;

                    try
                    {
                        imapFodler.Open(FolderAccess.ReadOnly);
                    }
                    catch (Exception)
                    {
                        logError.Add("Error: can't select imap folder '" + folder.Folder + "'");
                        return;
                    }

                    //IList<IMessageSummary> items = imapFodler.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size);
                    IList<IMessageSummary> items = imapFodler.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size);

                    DateTime dt = DateTime.Now;
                    long fileSize = 0;
                    MimeMessage msg;

                    long folderSize = items.Sum(o => o.Size ?? 0);

                    foreach (var item in items)
                    {

                        if (dc.MergeFolders)
                        {
                            var file = Directory.GetFiles(destFolder, item.UniqueId + "_*.eml");
                            if (file.Count() == 1)
                            {
                                logError.Add("Log: message id " + item.UniqueId + " already downloaded from folder '" + folder.Folder + "'");
                                downloadedEmails++;
                                folder.DownloadedItems++;
                                continue;
                            }
                            else { }
                        }

                        dt = DateTime.Now;
                        fileSize = 0;

                        try
                        {
                            msg = imapFodler.GetMessage(item.UniqueId);
                        }
                        catch
                        {
                            // Second attempt
                            try
                            {
                                msg = imapFodler.GetMessage(item.UniqueId);
                            }
                            catch (Exception ex)
                            {
                                // in the meanwhile a message has been deleted.. sometimes happens
                                logError.Add("Error: can't download message id " + item.UniqueId + " from folder '" + folder.Folder + "'");
                                continue;
                            }
                        }

                        ProgressMails++;

                        if (folder.Selected == false)
                            continue;

                        // msg not exsist
                        if (msg.From == null)
                        {
                            log.Add("Error: can't save message id " + item.UniqueId + " from folder '" + folder.Folder + "' because has no From field");
                            continue;
                        }

                        downloadedEmails++;
                        folder.DownloadedItems++;

                        messageIdSafeName = System.Text.RegularExpressions.Regex.Replace(msg.Headers["Message-ID"] + "", "[<>\\/]", "");

                        if (string.IsNullOrEmpty(messageIdSafeName))
                        {
                            messageIdSafeName = Guid.NewGuid().ToString();
                        }
                        else if (messageIdSafeName.Length > 250)
                        {
                            // i'll take the lst 250 characters
                            messageIdSafeName = messageIdSafeName.Substring(messageIdSafeName.Length - 250);
                        }
                        try
                        {
                            using (var fs = new FileStream(Path.Combine(destFolder, item.UniqueId + "_" + messageIdSafeName + ".eml"), FileMode.Create))
                            {
                                msg.WriteTo(fs);
                                fileSize = fs.Length;
                            }
                        }
                        catch (PathTooLongException)
                        {
                            logError.Add("Warning: message id " + item.UniqueId + " from folder '" + folder.Folder + "' will be saved with name '" + item.UniqueId + ".eml' because '" + item.UniqueId + "_" + messageIdSafeName + ".eml' is too long");
                            using (var fs = new FileStream(Path.Combine(destFolder, item.UniqueId + ".eml"), FileMode.Create))
                            {
                                msg.WriteTo(fs);
                                fileSize = fs.Length;
                            }
                        }

                        DownloadSpeed.Add(new Tuple<DateTime, double, long>(dt, DateTime.Now.Subtract(dt).TotalMilliseconds, fileSize));

                    }

                    folder.IsDownloading = false;

                    try
                    {
                        imapFodler.Close();
                    }
                    catch (MailKit.ServiceNotConnectedException)
                    {

                    }

                    log.Add("Folder: " + folder.Folder + "\t\t" + downloadedEmails + " emails");

                    totalMessagesDownloaded += downloadedEmails;

                }

            });

            log.Add("");

            log.Add("Total emails: " + totalMessagesDownloaded);

            DateTime dtEnd = DateTime.UtcNow;

            log.Add("");
            log.Add("Startd at " + dtStart.ToUniversalTime() + " UTC");
            log.Add("End at " + dtEnd.ToUniversalTime() + " UTC");

            log.Add("");

            dc.PartialPercent = 100;

            string lastDirName = dc.DestinationFolder.Split('\\', '/').Where(o => !string.IsNullOrEmpty(o)).Last();
            string superDir = dc.DestinationFolder.Substring(0, dc.DestinationFolder.Length - lastDirName.Length - 1);
            string zipFileName = Path.Combine(superDir, lastDirName + ".zip");

            if (File.Exists(zipFileName))
                File.Delete(zipFileName);

            dc.PartialPercent = 0;

            dc.Speed30sec = "";
            dc.SpeedTotal = "Creating archive..";

            ZipFile.CreateFromDirectory(dc.DestinationFolder, zipFileName, CompressionLevel.Fastest, true);

            log.Add("Export file : " + zipFileName);

            dc.Speed30sec = "";
            dc.SpeedTotal = "Calculating hash..";

            string md5 = CalculateMD5(zipFileName).Replace("-", "");
            string sha1 = CalculateSHA1(zipFileName).Replace("-", "");

            log.Add("MD5 : " + md5);
            log.Add("SHA1 : " + sha1);

            string logFileName = Path.Combine(superDir, lastDirName + ".log");

            if (File.Exists(logFileName))
                File.Delete(logFileName);

            File.WriteAllLines(logFileName, log.Union(logError));

            Directory.Delete(dc.DestinationFolder, true);

            dc.PartialPercent = 100;

            dc.Speed30sec = "DONE! ";
            dc.SpeedTotal = string.Format(" It took {0} ", DateTime.UtcNow.Subtract(dtStart));

            return totalMessagesDownloaded;

        }


        private static string CalculateSHA1(string filePath)
        {
            using (SHA1CryptoServiceProvider cryptoTransformSHA1 = new SHA1CryptoServiceProvider())
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open))
                {
                    return BitConverter.ToString(cryptoTransformSHA1.ComputeHash(file));
                }
            }
        }

        private static string CalculateMD5(string filePath)
        {
            using (MD5CryptoServiceProvider cryptoTransformSHA1 = new MD5CryptoServiceProvider())
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open))
                {
                    return BitConverter.ToString(cryptoTransformSHA1.ComputeHash(file));
                }
            }
        }

    }

}
