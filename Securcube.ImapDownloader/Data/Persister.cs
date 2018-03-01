using Microsoft.Win32;
using SecurCube.ImapDownloader.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Securcube.ImapDownloader.Data
{
    internal class Persister
    {

        const string RegKeyName = @"SOFTWARE\SecurCube\ImapDownloader\History";

        private static RegistryKey getRegKey( bool writable = false)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(RegKeyName, writable);
            if (key == null)
            {
                key= Registry.CurrentUser.CreateSubKey(RegKeyName, writable);
            }
            return key;
        }


        static internal IEnumerable<DataContext> GetConnectionHistory()
        {
            RegistryKey key = getRegKey();
            DataContext dc;
            foreach (var item in key.GetValueNames())
            {
                string subKey = key.GetValue(item) as string;

                if (!string.IsNullOrWhiteSpace(subKey))
                {
                    try
                    {
                        dc = Serializer.DeserializeFromString<DataContext>(subKey);
                    }
                    catch (Exception)
                    {
                        dc = null;
                    }
                    if (dc != null)
                        yield return dc;
                }
            }
        }

        static internal void UpdateConnectionHistory(DataContext dc)
        {
            RegistryKey key = getRegKey(true);
            foreach (var item in key.GetValueNames())
            {
                if (item == dc.HostName + ":" + dc.UserName)
                {
                    key.DeleteValue(item);
                }
            }

            key.SetValue(dc.HostName + ":" + dc.UserName, Serializer.SerializeToString(dc), RegistryValueKind.String);

        }

    }
}
