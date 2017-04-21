using MailKit;

namespace Securcube.ImapDownloader.Data
{
    class EmailFolder : MyNotifyPropertyChanged
    {

        private bool selected;
        public bool Selected
        {
            get { return selected; }
            set { SetField(ref selected, value, "Selected"); }
        }

        private int? downloadedItems;
        public int? DownloadedItems
        {
            get { return downloadedItems; }
            set { SetField(ref downloadedItems, value, "DownloadedItems"); }
        }

        public int Unread { get; internal set; }
        public string Folder { get; internal set; }
        public int Messages { get; internal set; }
        public uint NextUID { get; internal set; }

        public FolderAttributes Flags { get; internal set; }
        public string FlagsToString
        {
            get
            {
                return Flags.ToString();
            }
        }

    }
}
