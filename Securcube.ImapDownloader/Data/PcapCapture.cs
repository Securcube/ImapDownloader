using PcapDotNet.Core;
using System;
using System.Threading;

namespace Securcube.ImapDownloader.Data
{
    public class PcapCapture : IDisposable
    {

        PacketDumpFile dumpFile { get; set; }
        PacketCommunicator communicator { get; set; }
        LivePacketDevice CapturePcapDevice { get; set; }
        public string OutputFile { get; private set; }

        public bool IsCapturing { get; private set; }

        Thread capturingThread;

        public PcapCapture(LivePacketDevice CapturePcapDevice, string outputFile)
        {
            this.CapturePcapDevice = CapturePcapDevice;
            this.OutputFile = outputFile;
        }


        internal void StartCapture()
        {

            if (CapturePcapDevice == null)
                return;

            if (string.IsNullOrWhiteSpace(OutputFile))
                return;

            capturingThread = new Thread(_StartCapture);
            capturingThread.Start();

        }

        private void _StartCapture()
        {
            IsCapturing = true;
            // Open the device
            communicator =
                           CapturePcapDevice.Open(65536,                                  // portion of the packet to capture
                                                                                          // 65536 guarantees that the whole packet will be captured on all the link layers
                                               PacketDeviceOpenAttributes.Promiscuous,    // promiscuous mode
                                               1000);                                     // read timeout

            dumpFile = communicator.OpenDump(OutputFile);

            try
            {
                // start the capture
                communicator.ReceivePackets(0, dumpFile.Dump);
            }
            catch (Exception ex)
            {
                // .... why??
            }

        }

        internal void StopCapture()
        {
            if (capturingThread != null)
            {

                communicator.Break();

                dumpFile.Flush();
                dumpFile.Dispose();

                communicator.Dispose();

                capturingThread.Abort();
            }
            capturingThread = null;
            IsCapturing = false;
        }

        public void Dispose()
        {
            StopCapture();
        }
    }
}
