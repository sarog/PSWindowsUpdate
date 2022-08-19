namespace PSWindowsUpdate {
    public class WUSettings {
        public string ComputerName { get; set; }

        public int AcceptTrustedPublisherCerts { get; set; }

        public string WUServer { get; set; }

        public string WUStatusServer { get; set; }

        public int DetectionFrequencyEnabled { get; set; }

        public int DetectionFrequency { get; set; }

        public int NoAutoRebootWithLoggedOnUsers { get; set; }

        public int RebootRelaunchTimeoutEnabled { get; set; }

        public int RebootRelaunchTimeout { get; set; }

        public int IncludeRecommendedUpdates { get; set; }

        public int NoAutoUpdate { get; set; }

        public string AUOptions { get; set; }

        public string ScheduledInstallDay { get; set; }

        public int ScheduledInstallTime { get; set; }

        public int UseWUServer { get; set; }
    }
}