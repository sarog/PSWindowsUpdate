namespace PSWindowsUpdate
{
    /// <summary>Return error code description.</summary>
    public class WUApiCode
    {
        /// <summary>ErrorCode in Hex</summary>
        public string HResult { get; set; }

        /// <summary>Short message</summary>
        public string Message { get; set; }

        /// <summary>Full Description</summary>
        public string Description { get; set; }

        /// <summary>Type of WUApiCode</summary>
        public int CodeType { get; set; }
    }
}