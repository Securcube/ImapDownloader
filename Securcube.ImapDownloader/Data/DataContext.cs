using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurCube.ImapDownloader.Data
{
    class DataContext : MyNotifyPropertyChanged
    {

        private string hostName;
        public string HostName
        {
            get { return hostName; }
            set { SetField(ref hostName, value, "HostName"); }
        }

        private string userName;
        public string UserName
        {
            get { return userName; }
            set { SetField(ref userName, value, "UserName"); }
        }

        private string userPassword;
        public string UserPassword
        {
            get { return userPassword; }
            set { SetField(ref userPassword, value, "UserPassword"); }
        }

        private bool useSSL;
        public bool UseSSL
        {
            get { return useSSL; }
            set { SetField(ref useSSL, value, "UseSSL"); }
        }

        private int port;
        public int Port
        {
            get { return port; }
            set { SetField(ref port, value, "Port"); }
        }

        private string destinationFolder;
        public string DestinationFolder
        {
            get { return destinationFolder; }
            set { SetField(ref destinationFolder, value, "DestinationFolder"); }
        }


        private int concurrentThreads = 5;
        public int ConcurrentThreads
        {
            get { return concurrentThreads; }
            set { SetField(ref concurrentThreads, value, "ConcurrentThreads"); }
        }

        public List<Data.EmailFolder> EmailFolders { get; internal set; }


        private string speed30sec;
        public string Speed30sec
        {
            get { return speed30sec; }
            set { SetField(ref speed30sec, value, "Speed30sec"); }
        }

        private string speedTotal;
        public string SpeedTotal
        {
            get { return speedTotal; }
            set { SetField(ref speedTotal, value, "SpeedTotal"); }
        }

        private decimal partialPercent;

        public decimal PartialPercent
        {
            get { return partialPercent; }
            set { SetField(ref partialPercent, value, "PartialPercent"); }
        }
    }
}
