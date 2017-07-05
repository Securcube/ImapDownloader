using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media;

namespace SecurCube.ImapDownloader.Converters
{
    class ProgressForegroundConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values != null && values.Length == 3)
            {

                int? DownloadedItems = values[0] as int?;
                int? Messages = values[1] as int?;
                bool? isDownloading = values[2] as bool?;

                if (DownloadedItems != null && Messages != null)
                {
                    Brush foreground = Brushes.Green;

                    if (isDownloading == true)
                    {
                        if (DownloadedItems == Messages)
                            foreground = Brushes.LightGreen;
                        else if (DownloadedItems < Messages)
                            foreground = Brushes.YellowGreen;
                        else if (DownloadedItems > Messages)
                            foreground = Brushes.MediumPurple;
                    }
                    else
                    {
                        if (DownloadedItems == Messages)
                            foreground = Brushes.Green;
                        else if (DownloadedItems < Messages)
                            foreground = Brushes.Yellow;
                        else if (DownloadedItems > Messages)
                            foreground = Brushes.Purple;
                    }

                    return foreground;
                }

            }
            return null;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
