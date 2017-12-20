using System;
using System.Linq;
using System.Windows;
using System.IO;
using System.Threading.Tasks;
using System.Timers;
using Securcube.ImapDownloader.Data;
using System.Windows.Controls;

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

            var history = Persister.GetConnectionHistory();

            if (history != null && history.Any())
            {
                SelectHistory.Items.Add(new ComboBoxItem() { IsEnabled = false, IsSelected = true, Content = "Select Credentials" });
                foreach (var item in history)
                {
                    SelectHistory.Items.Add(new ComboBoxItem() { Content = string.Format("{0}\n{1}", item.HostName, item.UserName), DataContext = item });
                }
            }
            else
            {
                SelectHistory.Visibility = Visibility.Collapsed;
            }

        }

        private async void buttonTestParams_Click(object sender, RoutedEventArgs e)
        {

            var OrigText = buttonTestParams.Content.ToString();

            connectionParams.IsEnabled = false;
            try
            {
                buttonTestParams.Content = "Testing Parameters. Please wait!";
                dataGridFolders.ItemsSource = await Downloader.GetAllFoldersAsync(dc);
                Persister.UpdateConnectionHistory(dc);
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



        private async void ButtonStartDownload_Click(object sender, RoutedEventArgs e)
        {

            (sender as FrameworkElement).IsEnabled = false;
            ProgressBarBox.Visibility = Visibility.Visible;
            dataGridFolders.IsReadOnly = true;

            dc.MergeFolders = false;

            if (Directory.Exists(dc.DestinationFolder))
            {
                var result = MessageBox.Show(string.Format("The folder {0} already exsist.\nwould you resume the acquisition?\nClick 'Yes' to continue the acquisition. 'No' for start download all mails from the beginning", dc.DestinationFolder), "Merge mailbox", MessageBoxButton.YesNo);
                dc.MergeFolders = result == MessageBoxResult.Yes;
            }

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

                var tt = Downloader.getDownloadSpeed();

                double s1 = tt.Item1;
                double s2 = tt.Item2;

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

                if (Downloader.TotalMails != 0)
                {
                    dc.PartialPercent = (decimal)Downloader.ProgressMails / Downloader.TotalMails * 100;
                }
                else
                {
                    dc.PartialPercent = 0;
                }

            });
            myTimer.Interval = 5000;
            myTimer.Enabled = true;
            myTimer.Start();
            var result = await Downloader.DownloadMailsAsync(dc);
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

        private void SelectHistory_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {


            if (SelectHistory.SelectedIndex > 0 && SelectHistory.SelectedItem is ComboBoxItem  &&  (SelectHistory.SelectedItem as ComboBoxItem).DataContext is Data.DataContext)
            {
                var _dc = (SelectHistory.SelectedItem as ComboBoxItem).DataContext as Data.DataContext;

                dc.HostName = _dc.HostName;
                dc.UserName = _dc.UserName;
                dc.UserPassword = _dc.UserPassword;
                dc.UseSSL = _dc.UseSSL;
                dc.Port = _dc.Port;

                SelectHistory.SelectedIndex = 0;
            }

        }
    }
}
