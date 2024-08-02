namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for History objects.</para>
    /// </summary>
    public class History
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>Operationname</summary>
        public string Operationname { get; set; }

        /// <summary>Date</summary>
        public string Date { get; set; }

        /// <summary>KB</summary>
        public string KB { get; set; }

        /// <summary>Title</summary>
        public string Title { get; set; }
    }
}