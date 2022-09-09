using System;

namespace PSWindowsUpdate {
    public class RebootStatus {
        public string ComputerName { get; set; }

        public bool RebootRequired { get; set; }
        
        public DateTime? RebootScheduled { get; set; }
    }
}