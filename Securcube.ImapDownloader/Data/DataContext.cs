using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Securcube.ImapDownloader.Data
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


        public List<Data.EmailFolder> EmailFolders { get; internal set; }

        public char FolderSeparator { get; internal set; }




    }
}
