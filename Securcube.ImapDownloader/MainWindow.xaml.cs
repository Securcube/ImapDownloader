using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using System.IO.Compression;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Collections.Concurrent;
using System.Timers;

namespace SecurCube.ImapDownloader
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        Data.DataContext dc = new Data.DataContext();

        public MainWindow()
        {

            dc.UseSSL = true;
            dc.HostName = "imap.gmail.com";
            dc.UserName = "myUsername";
            dc.UserPassword = "myPassword";
            dc.Port = 993;

            this.DataContext = dc;
            InitializeComponent();
        }


        private char getSeparator(IList<IMailFolder> folders)
        {
            // default folder separator
            var fodlerSeparator = '/';
            var separators = new List<char>();

            for (int i = 1; i < folders.Count() - 1; i++)
            {
                if (folders[i].Name.StartsWith(folders[i - 1].Name))
                {
                    separators.Add(folders[i].Name[folders[i - 1].Name.Length]);
                }
            }

            if (separators.Any())
            {
                fodlerSeparator = separators.GroupBy(o => o).OrderBy(o => o.Count()).FirstOrDefault().Key;
            }

            return fodlerSeparator;

        }


        private async void buttonTestParams_Click(object sender, RoutedEventArgs e)
        {

            var OrigText = buttonTestParams.Content.ToString();

            connectionParams.IsEnabled = false;
            try
            {
                buttonTestParams.Content = "Testing Parameters. Please wait!";
                dataGridFolders.ItemsSource = await Task.Run(() => readFolders(dc));
                buttonTestParams.Visibility = Visibility.Collapsed;
                GridBottom.IsEnabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Not connected!\n" + ex.Message);
                buttonTestParams.Content = OrigText;
                connectionParams.IsEnabled = true;
            }
        }

        private static List<Data.EmailFolder> readFolders(Data.DataContext dc)
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

                var myFodlers = new List<Data.EmailFolder>();

                var folders = client.GetFolders(client.PersonalNamespaces[0]);

                foreach (var folder in folders)
                {
                    try
                    {
                        folder.Open(FolderAccess.ReadOnly);
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    myFodlers.Add(new Data.EmailFolder()
                    {
                        Folder = folder.FullName,
                        Messages = folder.Count,
                        NextUID = folder.UidNext != null ? folder.UidNext.Value.Id : 0,
                        Selected = !folder.Attributes.HasFlag(FolderAttributes.All),
                        Flags = folder.Attributes,
                        Unread = folder.Unread,
                    });

                    folder.Close();

                }

                dc.EmailFolders = myFodlers;

                return myFodlers;

            }

        }

        private async void ButtonStartDownload_Click(object sender, RoutedEventArgs e)
        {

            (sender as FrameworkElement).IsEnabled = false;
            ProgressBarBox.Visibility = Visibility.Visible;
            dataGridFolders.IsReadOnly = true;

            try
            {
                var result = await DownloadMail();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);

                (sender as FrameworkElement).IsEnabled = true;
                ProgressBarBox.Visibility = Visibility.Collapsed;
                dataGridFolders.IsReadOnly = false;

                //throw ex;
            }

        }


        private async Task<long> DownloadMail()
        {

            var lastQueueClean = DateTime.Now;

            var myTimer = new Timer();
            myTimer.Elapsed += new ElapsedEventHandler((s, e) =>
            {

                //var cleanBeforeDateTime = DateTime.Now.AddMinutes(-2);

                //if (lastQueueClean < cleanBeforeDateTime)
                //{
                //    var itemsToPurge = DownloadSpeed.Where(o=> o.Item1 <= cleanBeforeDateTime);

                //    DownloadSpeed.r

                //    var remove = DownloadSpeed.

                //}


                double s1 = 0, s2 = 0;

                // Check speed last 30 seconds
                var dt = DateTime.Now.AddSeconds(-30);
                var items = DownloadSpeed.Where(o => o.Item1 > dt);
                var downloadedData = items.Sum(o => o.Item3);
                var downloadedTime = items.Sum(o => o.Item2);
                if (downloadedTime != 0)
                {
                    s1 = downloadedData / downloadedTime;
                }

                // check spped from beginning
                downloadedData = DownloadSpeed.Sum(o => o.Item3);
                downloadedTime = DownloadSpeed.Sum(o => o.Item2);
                if (downloadedTime != 0)
                {
                    s2 = downloadedData / downloadedTime;
                }

                if (s1 > 0)
                {
                    dc.Speed30sec = SizeSuffix(s1) + "/s";

                    if (s1 != s2)
                    {
                        dc.SpeedTotal = " ( AVG " + SizeSuffix(s2) + "/s)";
                    }
                    else
                    {
                        dc.SpeedTotal = "";
                    }
                }

                if (TotalMails != 0)
                {
                    dc.PartialPercent = (decimal)ProgressMails / TotalMails * 100;
                }
                else
                {
                    dc.PartialPercent = 0;
                }

            });
            myTimer.Interval = 3000;
            myTimer.Enabled = true;
            myTimer.Start();
            var result = await Task.Run(() => _DownloadMail());
            myTimer.Stop();
            return result;
        }

        static readonly string[] SizeSuffixes = { "bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB" };
        static string SizeSuffix(double value, int decimalPlaces = 1)
        {
            if (value < 0) { return "-" + SizeSuffix(-value); }
            if (value == 0) { return "0.0 bytes"; }

            // mag is 0 for bytes, 1 for KB, 2, for MB, etc.
            int mag = (int)Math.Log(value, 1024);

            // 1L << (mag * 10) == 2 ^ (10 * mag) 
            // [i.e. the number of bytes in the unit corresponding to mag]
            decimal adjustedSize = (decimal)value / (1L << (mag * 10));

            // make adjustment when the value is large enough that
            // it would round up to 1000 or more
            if (Math.Round(adjustedSize, decimalPlaces) >= 1000)
            {
                mag += 1;
                adjustedSize /= 1024;
            }

            return string.Format("{0:n" + decimalPlaces + "} {1}",
                adjustedSize,
                SizeSuffixes[mag]);
        }



        // Check connection download speed: first param = 
        ConcurrentBag<Tuple<DateTime, double, long>> DownloadSpeed = new ConcurrentBag<Tuple<DateTime, double, long>>();

        int TotalMails = 0;
        int ProgressMails = 0;

        private long _DownloadMail()
        {

            List<string> log = new List<string>();

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

            TotalMails = dc.EmailFolders.Sum(o => o.Messages);

            Parallel.ForEach(dc.EmailFolders, new ParallelOptions() { MaxDegreeOfParallelism = dc.ConcurrentThreads }, (folder) =>
            {

                if (folder.Selected == false)
                    return;

                // The default port for IMAP over SSL is 993.
                using (ImapClient client = new ImapClient())
                {

                    client.ServerCertificateValidationCallback = (s, c, h, ee) => true;

                    client.Connect(dc.HostName, dc.Port, dc.UseSSL);

                    // Note: since we don't have an OAuth2 token, disable
                    // the XOAUTH2 authentication mechanism.
                    client.AuthenticationMechanisms.Remove("XOAUTH2");

                    client.Authenticate(dc.UserName, dc.UserPassword);

                    var imapFodler = client.GetFolder(folder.Folder);

                    folder.IsDownloading = true;

                    folder.DownloadedItems = 0;

                    string destFolder = Path.Combine(dc.DestinationFolder, folder.Folder.Replace(imapFodler.DirectorySeparator, '\\'));

                    // If the folder already exsist I have do delete it
                    if (Directory.Exists(destFolder))
                        Directory.Delete(destFolder, true);

                    Directory.CreateDirectory(destFolder);

                    string messageIdSafeName = "";

                    int downloadedEmails = 0;

                    try
                    {
                        imapFodler.Open(FolderAccess.ReadOnly);
                    }
                    catch (Exception)
                    {
                        log.Add("Error: can't select imap folder '" + folder.Folder + "'");
                        return;
                    }

                    IList<IMessageSummary> items = imapFodler.Fetch(0, -1, MessageSummaryItems.UniqueId);

                    DateTime dt = DateTime.Now;
                    long fileSize = 0;
                    MimeMessage msg;

                    foreach (var item in items)
                    {

                        try
                        {
                            msg = imapFodler.GetMessage(item.UniqueId);
                        }
                        catch (Exception ex)
                        {
                            // in the meanwhile a message has been deleted.. sometimes happens
                            log.Add("Error: can't download message with id " + item.UniqueId + " from folder '" + folder.Folder + "'");
                            continue;
                        }

                        ProgressMails++;

                        if (folder.Selected == false)
                            continue;

                        // msg not exsist
                        if (msg.From == null)
                        {
                            log.Add("Error: can't save message with id " + item.UniqueId + " from folder '" + folder.Folder + "' because has no From field");
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

                        dt = DateTime.Now;
                        fileSize = 0;
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
                            log.Add("Warning: message with id " + item.UniqueId + " from folder '" + folder.Folder + "' will be saved with name '" + item.UniqueId + ".eml' because '" + item.UniqueId + "_" + messageIdSafeName + ".eml' is too long");
                            using (var fs = new FileStream(Path.Combine(destFolder, item.UniqueId + ".eml"), FileMode.Create))
                            {
                                msg.WriteTo(fs);
                                fileSize = fs.Length;
                            }
                        }

                        DownloadSpeed.Add(new Tuple<DateTime, double, long>(dt, DateTime.Now.Subtract(dt).TotalMilliseconds, fileSize));

                    }

                    folder.IsDownloading = false;

                    imapFodler.Close();

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

            string lastDirName = dc.DestinationFolder.Split('\\', '/').Where(o => !string.IsNullOrEmpty(o)).Last();
            string superDir = dc.DestinationFolder.Substring(0, dc.DestinationFolder.Length - lastDirName.Length - 1);
            string zipFileName = Path.Combine(superDir, lastDirName + ".zip");

            if (File.Exists(zipFileName))
                File.Delete(zipFileName);

            ZipFile.CreateFromDirectory(dc.DestinationFolder, zipFileName, CompressionLevel.Fastest, true);

            log.Add("Export file : " + zipFileName);

            string md5 = CalculateMD5(zipFileName).Replace("-", "");
            string sha1 = CalculateSHA1(zipFileName).Replace("-", "");

            log.Add("MD5 : " + md5);
            log.Add("SHA1 : " + sha1);

            string logFileName = Path.Combine(superDir, lastDirName + ".log");

            if (File.Exists(logFileName))
                File.Delete(logFileName);

            File.WriteAllLines(logFileName, log);

            Directory.Delete(dc.DestinationFolder, true);

            return totalMessagesDownloaded;

        }


        public static string CalculateSHA1(string filePath)
        {
            using (SHA1CryptoServiceProvider cryptoTransformSHA1 = new SHA1CryptoServiceProvider())
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open))
                {
                    return BitConverter.ToString(cryptoTransformSHA1.ComputeHash(file));
                }
            }
        }


        public static string CalculateMD5(string filePath)
        {
            using (MD5CryptoServiceProvider cryptoTransformSHA1 = new MD5CryptoServiceProvider())
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open))
                {
                    return BitConverter.ToString(cryptoTransformSHA1.ComputeHash(file));
                }
            }
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

            if (e.AddedItems == null && e.AddedItems.Count > 0)
                return;

            if (e.AddedItems[0] as FrameworkElement == null)
                return;

            var _dc = (e.AddedItems[0] as FrameworkElement).DataContext as string;

            if (string.IsNullOrEmpty(_dc))
                return;

            var arr = _dc.Split('|');

            if (arr.Length == 3)
            {
                dc.HostName = arr[0];
                dc.Port = int.Parse(arr[1]);
                dc.UseSSL = bool.Parse(arr[2]);
            }




        }
    }
}
