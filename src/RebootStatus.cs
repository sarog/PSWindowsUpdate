using System;

namespace PSWindowsUpdate {
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for RebootStatus objects.</para>
    /// </summary>
    public class RebootStatus {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        public bool RebootRequired { get; set; }

        public DateTime? RebootScheduled { get; set; }
    }
}