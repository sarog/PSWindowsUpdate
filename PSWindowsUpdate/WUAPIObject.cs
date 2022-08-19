using System;
using System.Management.Automation;

namespace PSWindowsUpdate {
    public class WUAPIObject {
        public string ComputerName { get; set; }

        public bool Status { get; set; }

        public object Object { get; set; }

        public ErrorRecord Error { get; set; }

        public Exception Exception { get; set; }

        public string Mode { get; set; }
    }
}