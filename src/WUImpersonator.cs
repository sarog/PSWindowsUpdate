using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security.Principal;

namespace PSWindowsUpdate {
    /// <summary>
    /// Impersonation of a user. Allows to execute code under another user context.
    /// Please note that the account that instantiates the Impersonator class
    /// needs to have the 'Act as part of operating system' privilege set.
    /// </summary>
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public class WUImpersonator : IDisposable {
        private WindowsImpersonationContext impersonationContext;

        /// <summary>
        /// Initializes a new instance of the <see cref="T:PSWindowsUpdate.WUImpersonator" /> class and
        /// impersonates as a built in service account.
        /// </summary>
        /// <param name="builtinUser">The built in user to impersonate - either
        /// Local Service or Network Service. These users can only be impersonated
        /// by code running as System.</param>
        public WUImpersonator(BuiltinUser builtinUser)
            : this(string.Empty, "NT AUTHORITY", string.Empty, LogonSessionType.Service, builtinUser: BuiltinUser.LocalService) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="T:PSWindowsUpdate.WUImpersonator" /> class and
        /// impersonates with the specified credentials.
        /// </summary>
        /// <param name="username">his is the name of the user account to log on
        /// to. If you use the user principal name (UPN) format,
        /// user@DNS_domain_name, the lpszDomain parameter must be <c>null</c>.</param>
        /// <param name="domain">The name of the domain or server whose account
        /// database contains the lpszUsername account. If this parameter is
        /// <c>null</c>, the user name must be specified in UPN format. If this
        /// parameter is ".", the function validates the account by using only the
        /// local account database.</param>
        /// <param name="password">The plaintext password for the user account.</param>
        public WUImpersonator(string username, string domain, string password)
            : this(username, domain, password, LogonSessionType.Interactive) { }

        /// <summary>
        /// Constructor. Starts the impersonation with the given credentials.
        /// Please note that the account that instantiates the Impersonator class
        /// needs to have the 'Act as part of operating system' privilege set.
        /// </summary>
        /// <param name="userName">The name of the user to act as.</param>
        /// <param name="domainName">The domain name of the user to act as.</param>
        /// <param name="password">The password of the user to act as.</param>
        /// <param name="logonType">Logon Session Type</param>
        /// <param name="logonProvider">Logon Provider</param>
        /// <param name="builtinUser">Builtin User</param>
        public WUImpersonator(
            string userName,
            string domainName,
            string password,
            LogonSessionType logonType = LogonSessionType.Interactive,
            LogonProvider logonProvider = LogonProvider.Default,
            BuiltinUser builtinUser = BuiltinUser.None) {
            switch (builtinUser) {
                case BuiltinUser.None:
                    if (string.IsNullOrEmpty(userName)) {
                        return;
                    }

                    break;
                case BuiltinUser.LocalService:
                    userName = "LOCAL SERVICE";
                    break;
                case BuiltinUser.NetworkService:
                    userName = "NETWORK SERVICE";
                    break;
            }

            ImpersonateValidUser(userName, domainName, password, logonType, logonProvider);
        }

        /// <summary>Destructor</summary>
        public void Dispose() {
            UndoImpersonation();
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int LogonUser(
            string lpszUserName,
            string lpszDomain,
            string lpszPassword,
            int dwLogonType,
            int dwLogonProvider,
            out IntPtr phToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int DuplicateToken(
            IntPtr hToken,
            int impersonationLevel,
            out IntPtr hNewToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool RevertToSelf();

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern bool CloseHandle(IntPtr handle);

        /// <summary>Does the actual impersonation.</summary>
        /// <param name="userName">The name of the user to act as.</param>
        /// <param name="domainName">The domain name of the user to act as.</param>
        /// <param name="password">The password of the user to act as.</param>
        /// <param name="logonType">Logon Session Type</param>
        /// <param name="logonProvider">Logon Provider</param>
        private void ImpersonateValidUser(
            string userName,
            string domainName,
            string password,
            LogonSessionType logonType,
            LogonProvider logonProvider) {
            var phToken = IntPtr.Zero;
            var hNewToken = IntPtr.Zero;
            try {
                if (!RevertToSelf()) {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                if (LogonUser(userName, domainName, password, (int)logonType, (int)logonProvider, out phToken) == 0) {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                if (DuplicateToken(phToken, 3, out hNewToken) == 0) {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                impersonationContext = new WindowsIdentity(hNewToken).Impersonate();
            } finally {
                if (phToken != IntPtr.Zero) {
                    CloseHandle(phToken);
                }

                if (hNewToken != IntPtr.Zero) {
                    CloseHandle(hNewToken);
                }
            }
        }

        /// <summary>Reverts the impersonation.</summary>
        private void UndoImpersonation() {
            if (impersonationContext == null) {
                return;
            }

            impersonationContext.Undo();
        }

        /// <summary>Specifies the type of login session used.</summary>
        public enum LogonSessionType {
            /// <summary>
            /// Intended for users who are interactively using the machine, such as a user being logged on by a terminal server, remote shell, or similar process.
            /// </summary>
            Interactive = 2,

            /// <summary>
            /// Intended for high-performance servers to authenticate clear text passwords. LogonUser does not cache credentials for this logon type.
            /// </summary>
            Network = 3,

            /// <summary>
            /// Intended for batch servers, where processes can be executed on behalf of a user without their direct intervention; or for higher performance servers that process many clear-text authentication attempts at a time, such as mail or web servers. LogonUser does not cache credentials for this logon type.
            /// </summary>
            Batch = 4,

            /// <summary>
            /// Indicates a service-type logon. The account provided must have the service privilege enabled.
            /// </summary>
            Service = 5,

            /// <summary>Indicates a proxy-type logon.</summary>
            Proxy = 6,

            /// <summary>
            /// This logon type is intended for GINA DLLs logging on users who are interactively using the machine. This logon type allows a unique audit record to be generated that shows when the workstation was unlocked.
            /// </summary>
            Unlock = 7,

            /// <summary>
            /// Preserves the name and password in the authentication packages, allowing the server to make connections to other network servers while impersonating the client. This allows a server to accept clear text credentials from a client, call LogonUser, verify that the user can access the system across the network, and still communicate with other servers.
            /// </summary>
            NetworkCleartext = 8,

            /// <summary>
            /// Allows the caller to clone its current token and specify new credentials for outbound connections. The new logon session has the same local identify, but uses different credentials for other network connections.
            /// </summary>
            NewCredentials = 9,

            /// <summary>
            /// Terminal Services session that is both remote and interactive.
            /// </summary>
            RemoteInteractive = 10, // 0x0000000A

            /// <summary>
            /// Attempt cached credentials without accessing the network.
            /// </summary>
            CachedInteractive = 11, // 0x0000000B

            /// <summary>
            /// Same as RemoteInteractive. This is used for internal auditing.
            /// </summary>
            CachedRemoteInteractive = 12, // 0x0000000C

            /// <summary>Workstation logon.</summary>
            CachedUnlock = 13 // 0x0000000D
        }

        /// <summary>Specifies the logon provider.</summary>
        public enum LogonProvider {
            /// <summary>
            /// Use the standard logon provider for the system. The default security provider is negotiate, unless you pass NULL for the domain name and the user name is not in UPN format. In this case, the default provider is NTLM.
            /// </summary>
            Default,

            /// <summary>
            /// </summary>
            WinNT35,

            /// <summary>Use the NTLM logon provider.</summary>
            WinNT40,

            /// <summary>Use the negotiate logon provider.</summary>
            WinNT50
        }

        /// <summary>Specifies the impersonation level.</summary>
        public enum ImpersonationLevel {
            /// <summary>
            /// The client is anonymous to the server. The server process can impersonate the client, but the impersonation token does not contain any information about the client. This level is only supported over the local interprocess communication transport. All other transports silently promote this level to identify.
            /// </summary>
            Anonymous,

            /// <summary>
            /// The system default level. The server can obtain the client's identity, and the server can impersonate the client to do ACL checks.
            /// </summary>
            Identify,

            /// <summary>
            /// The server can impersonate the client's security context while acting on behalf of the client. The server can access local resources as the client. If the server is local, it can access network resources as the client. If the server is remote, it can access only resources that are on the same computer as the server.
            /// </summary>
            Impersonate,

            /// <summary>
            /// The most powerful impersonation level. When this level is selected, the server (whether local or remote) can impersonate the client's security context while acting on behalf of the client. During impersonation, the client's credentials (both local and network) can be passed to any number of computers.
            /// </summary>
            Delegate
        }

        /// <summary>Specifies the BuiltinUser type.</summary>
        public enum BuiltinUser {
            /// <summary>None.</summary>
            None,

            /// <summary>Local Service.</summary>
            LocalService,

            /// <summary>Network Service.</summary>
            NetworkService
        }
    }
}