namespace PSWindowsUpdate
{
    /// <summary>
    /// OutputType
    /// <para type="description">OutputType for AgentInfo objects.</para>
    /// </summary>
    public class AgentInfo
    {
        /// <summary>ComputerName</summary>
        public string ComputerName { get; set; }

        /// <summary>ApiVersion</summary>
        public string ApiVersion { get; set; }

        /// <summary>WuapiDll</summary>
        public string WuapiDll { get; set; }
    }
}