using System;
using System.Management.Automation;

namespace PSWindowsUpdate {
    /// <summary>Return Windows Update API COM object.</summary>
    public class WUAPIObject {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>Status</summary>
        public bool Status { get; set; }

        /// <summary>Object</summary>
        public object Object { get; set; }

        /// <summary>Error</summary>
        public ErrorRecord Error { get; set; }

        /// <summary>Exception</summary>
        public Exception Exception { get; set; }

        /// <summary>Mode</summary>
        public string Mode { get; set; }
    }
}