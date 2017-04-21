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

namespace Securcube.ImapDownloader
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

            try
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

                    connectionParams.IsEnabled = false;
                    buttonTestParams.Visibility = Visibility.Collapsed;
                    gridDownloadProcess.Visibility = Visibility.Visible;
                    dc.DestinationFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\ExportImap\\" + dc.UserName + "\\";

                    var myFodlers = new List<Data.EmailFolder>();

                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);

                    var fodlerSeparator = getSeparator(folders);

                    foreach (var folder in folders)
                    {
                        folder.Open(FolderAccess.ReadOnly);

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
                    dc.FolderSeparator = fodlerSeparator;

                    dataGridFolders.ItemsSource = myFodlers;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Not connected!\n" + ex.Message);

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
                throw ex;
            }

            dataGridFolders.IsReadOnly = false;
            ProgressBarBox.Visibility = Visibility.Collapsed;
            (sender as FrameworkElement).IsEnabled = true;

        }


        private async Task<long> DownloadMail()
        {
            return await Task.Run(() => _DownloadMail());
        }

        private long _DownloadMail()
        {

            var log = new List<string>();

            // The default port for IMAP over SSL is 993.
            using (ImapClient client = new ImapClient())
            {

                client.ServerCertificateValidationCallback = (s, c, h, ee) => true;

                client.Connect(dc.HostName, dc.Port, dc.UseSSL);


                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(dc.UserName, dc.UserPassword);


                var dtStart = DateTime.UtcNow;

                log.Add("");
                log.Add("LOGIN PARAMS");
                log.Add("Host name: " + dc.HostName + " (resolved ip = " + System.Net.Dns.GetHostEntry(dc.HostName).AddressList.FirstOrDefault() + ") ");
                log.Add("Port: " + dc.Port);
                log.Add("UseSSL: " + dc.UseSSL.ToString());
                log.Add("User name: " + dc.UserName);
                log.Add("User password: " + dc.UserPassword);
                log.Add("");


                long totalMessagesDownloaded = 0;

                var folders = client.GetFolders(client.PersonalNamespaces[0]);

                foreach (var imapFodler in folders)
                {

                    var folder = dc.EmailFolders.Where(o => o.Folder == imapFodler.FullName).First();

                    if (folder.Selected == false)
                        continue;

                    folder.DownloadedItems = 0;

                    var destFolder = Path.Combine(dc.DestinationFolder, folder.Folder.Replace(dc.FolderSeparator, '\\'));

                    // If the folder already exsist I have do delete it
                    if (Directory.Exists(destFolder))
                        Directory.Delete(destFolder, true);

                    Directory.CreateDirectory(destFolder);


                    string messageIdSafeName = "";

                    var downloadedEmails = 0;


                    imapFodler.Open(FolderAccess.ReadOnly);

                    foreach (var msg in imapFodler)
                    {


                        //try
                        //{
                        // indexes of AE.Net.Mail are in base 0. They will be increased by one

                        if (folder.Selected == false)
                            continue;

                        // msg not exsist
                        if (msg.From == null && (msg.To == null || msg.To.Count == 0))
                            continue;

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
                            using (var fs = new FileStream(Path.Combine(destFolder, folder.DownloadedItems + "_" + messageIdSafeName + ".eml"), FileMode.Create))
                            {
                                msg.WriteTo(fs);
                            }
                        }
                        catch (PathTooLongException)
                        {
                            using (var fs = new FileStream(Path.Combine(destFolder, folder.DownloadedItems + ".eml"), FileMode.Create))
                            {
                                msg.WriteTo(fs);
                            }
                        }

                        //}
                        //catch (Exception ex)
                        //{
                        //    log.Add("Excaption : " + fodler.Folder + " " + downloadedEmails + " emails downloaded, Error: " + ex.Message);
                        //}


                    }

                    imapFodler.Close();



                    log.Add("Folder: " + folder.Folder + "\t\t" + downloadedEmails + " emails");

                    totalMessagesDownloaded += downloadedEmails;

                }

                log.Add("");

                log.Add("Total emails: " + totalMessagesDownloaded);

                log.Add("");
                log.Add("Startd at " + dtStart.ToUniversalTime() + " UTC");
                log.Add("End at " + dtStart.ToUniversalTime() + " UTC");

                log.Add("");

                var dtEnd = DateTime.UtcNow;

                var lastDirName = dc.DestinationFolder.Split('\\', '/').Where(o => !string.IsNullOrEmpty(o)).Last();
                var superDir = dc.DestinationFolder.Substring(0, dc.DestinationFolder.Length - lastDirName.Length - 1);
                var zipFileName = Path.Combine(superDir, lastDirName + ".zip");

                if (File.Exists(zipFileName))
                    File.Delete(zipFileName);

                ZipFile.CreateFromDirectory(dc.DestinationFolder, zipFileName, CompressionLevel.Fastest, true);

                log.Add("Export file : " + zipFileName);

                var md5 = CalculateMD5(zipFileName).Replace("-", "");
                var sha1 = CalculateSHA1(zipFileName).Replace("-", "");

                log.Add("MD5 : " + md5);
                log.Add("SHA1 : " + sha1);

                var logFileName = Path.Combine(superDir, lastDirName + ".log");

                if (File.Exists(logFileName))
                    File.Delete(logFileName);

                File.WriteAllLines(logFileName, log);

                Directory.Delete(dc.DestinationFolder, true);

                return totalMessagesDownloaded;

            }

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

    }
}
