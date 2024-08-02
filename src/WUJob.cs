namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for WUJob.</para>
    /// </summary>
    public class WUJob
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>Name</summary>
        public string Name { get; set; }

        /// <summary>Action</summary>
        public string Action { get; set; }
    }
}