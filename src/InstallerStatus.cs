namespace PSWindowsUpdate {
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for InstallerStatus objects.</para>
    /// </summary>
    public class InstallerStatus {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>IsBusy</summary>
        public bool IsBusy { get; set; }
    }
}