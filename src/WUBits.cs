using BITSReference4_0;
using System;
using System.Runtime.InteropServices;

namespace PSWindowsUpdate
{
    internal class WUBits
    {
        /// <summary>Download files from url using BITS.</summary>
        /// <param name="Source">Source URL</param>
        /// <param name="Destination">Destination path</param>
        public string StartBitsTransfer(string Source, string Destination)
        {
            var backgroundCopyManager4_ =
                (BackgroundCopyManager4_0)Activator.CreateInstance(
                    Marshal.GetTypeFromCLSID(new Guid("BB6DF56B-CACE-11DC-9992-0019B93A3A84")));
            backgroundCopyManager4_.CreateJob("Quick download", BG_JOB_TYPE.BG_JOB_TYPE_DOWNLOAD, out var _, out var ppJob);
            ppJob.AddFile(Source, Destination);
            ppJob.Resume();
            var flag = false;
            var result = "";
            while (!flag)
            {
                ppJob.GetState(out var pVal);
                switch (pVal)
                {
                    case BG_JOB_STATE.BG_JOB_STATE_ERROR:
                        result = "ERROR";
                        flag = true;
                        break;
                    case BG_JOB_STATE.BG_JOB_STATE_TRANSFERRED:
                        ppJob.Complete();
                        result = "TRANSFERRED";
                        break;
                    case BG_JOB_STATE.BG_JOB_STATE_CANCELLED:
                        result = "CANCELLED";
                        break;
                    case BG_JOB_STATE.BG_JOB_STATE_ACKNOWLEDGED:
                        result = "ACKNOWLEDGED";
                        flag = true;
                        break;
                }
            }

            return result;
        }
    }
}