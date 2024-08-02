namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for WUSettings objects.</para>
    /// </summary>
    public class WUSettings
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>AcceptTrustedPublisherCerts</summary>
        public int AcceptTrustedPublisherCerts { get; set; }

        /// <summary>WUServer</summary>
        public string WUServer { get; set; }

        /// <summary>WUStatusServer</summary>
        public string WUStatusServer { get; set; }

        /// <summary>DetectionFrequencyEnabled</summary>
        public int DetectionFrequencyEnabled { get; set; }

        /// <summary>DetectionFrequency</summary>
        public int DetectionFrequency { get; set; }

        /// <summary>NoAutoRebootWithLoggedOnUsers</summary>
        public int NoAutoRebootWithLoggedOnUsers { get; set; }

        /// <summary>RebootRelaunchTimeoutEnabled</summary>
        public int RebootRelaunchTimeoutEnabled { get; set; }

        /// <summary>RebootRelaunchTimeout</summary>
        public int RebootRelaunchTimeout { get; set; }

        /// <summary>IncludeRecommendedUpdates</summary>
        public int IncludeRecommendedUpdates { get; set; }

        /// <summary>NoAutoUpdate</summary>
        public int NoAutoUpdate { get; set; }

        /// <summary>AUOptions</summary>
        public string AUOptions { get; set; }

        /// <summary>ScheduledInstallDay</summary>
        public string ScheduledInstallDay { get; set; }

        /// <summary>ScheduledInstallTime</summary>
        public int ScheduledInstallTime { get; set; }

        /// <summary>UseWUServer</summary>
        public int UseWUServer { get; set; }
    }
}