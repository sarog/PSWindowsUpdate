namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for credentials from Credential Manager.</para>
    /// </summary>
    public class PSCredMan
    {
        /// <summary>Username</summary>
        public string Username { get; set; }

        /// <summary>Password</summary>
        public string Password { get; set; }

        /// <summary>Domain</summary>
        public string Domain { get; set; }
    }
}