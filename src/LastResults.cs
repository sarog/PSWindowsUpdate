using System;

namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for LastResult objects.</para>
    /// </summary>
    public class LastResults
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>LastSearchSuccessDate</summary>
        public DateTime LastSearchSuccessDate { get; set; }

        /// <summary>LastInstallationSuccessDate</summary>
        public DateTime LastInstallationSuccessDate { get; set; }
    }
}