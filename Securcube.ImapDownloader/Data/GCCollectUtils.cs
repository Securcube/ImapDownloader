using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Securcube.ImapDownloader.Data
{
    class GCCollectUtils
    {

        /// <summary>
        /// Last GC.Collect() call
        /// </summary>
        static DateTime LastGCMemoryClean = DateTime.Now;

        /// <summary>
        /// Memory usage limit: 250MB
        /// </summary>
        const long MemSizeLimit = 250 * 1024 * 1024;

        /// <summary>
        /// check if memory usage is too big
        /// </summary>
        internal static void CheckAndFreeMemory()
        {
            try
            {
                // Don't stress GC
                if (DateTime.Now.Subtract(LastGCMemoryClean).TotalSeconds < 10)
                {
                    return;
                }

                long MemSize = GC.GetTotalMemory(true);
                if (MemSize > MemSizeLimit)
                {
                    LastGCMemoryClean = DateTime.Now;
                    GC.Collect();
                }
            }
            catch
            {
                // No problem if exceptions...            
            }

        }

    }

}
