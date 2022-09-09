using System;

namespace PSWindowsUpdate {
    public class LastResults {
        public string ComputerName { get; set; }

        public DateTime LastSearchSuccessDate { get; set; }

        public DateTime LastInstallationSuccessDate { get; set; }
    }
}