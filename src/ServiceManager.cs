namespace PSWindowsUpdate {
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for ServiceManager objects.</para>
    /// </summary>
    public class ServiceManager {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>Name</summary>
        public string Name { get; set; }

        /// <summary>ServiceID</summary>
        public string ServiceID { get; set; }

        /// <summary>IsManaged</summary>
        public string IsManaged { get; set; }

        /// <summary>IsDefaultAUService</summary>
        public string IsDefaultAUService { get; set; }
    }
}