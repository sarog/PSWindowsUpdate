namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for WindowsUpdate objects.</para>
    /// </summary>
    public class WindowsUpdate
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>Status</summary>
        public string Status { get; set; }

        /// <summary>KB</summary>
        public string KB { get; set; }

        /// <summary>Size</summary>
        public string Size { get; set; }

        /// <summary>Title</summary>
        public string Title { get; set; }
    }
}