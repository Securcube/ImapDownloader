using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace Securcube.ImapDownloader.Data
{

    internal class Serializer
    {

        internal static TData DeserializeFromString<TData>(string settings)
        {
            byte[] b = Convert.FromBase64String(settings);
            using (var stream = new MemoryStream(b))
            {
                var formatter = new BinaryFormatter();
                formatter.FilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Low;
                stream.Seek(0, SeekOrigin.Begin);
                return (TData)formatter.Deserialize(stream);
            }
        }

        internal static string SerializeToString<TData>(TData settings)
        {
            using (var stream = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.FilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Low;
                formatter.Serialize(stream, settings);
                stream.Flush();
                stream.Position = 0;
                return Convert.ToBase64String(stream.ToArray());
            }
        }


    }
}
