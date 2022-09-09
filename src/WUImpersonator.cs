using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security.Principal;

namespace PSWindowsUpdate {
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public class WUImpersonator : IDisposable {
        private WindowsImpersonationContext impersonationContext;

        public WUImpersonator(BuiltinUser builtinUser)
            : this(string.Empty, "NT AUTHORITY", string.Empty, LogonSessionType.Service, builtinUser: BuiltinUser.LocalService) { }

        public WUImpersonator(string username, string domain, string password)
            : this(username, domain, password, LogonSessionType.Interactive) { }

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

        private void UndoImpersonation() {
            if (impersonationContext == null) {
                return;
            }

            impersonationContext.Undo();
        }

        public enum LogonSessionType {
            Interactive = 2,
            Network = 3,
            Batch = 4,
            Service = 5,
            Proxy = 6,
            Unlock = 7,
            NetworkCleartext = 8,
            NewCredentials = 9,
            RemoteInteractive = 10,       // 0x0000000A
            CachedInteractive = 11,       // 0x0000000B
            CachedRemoteInteractive = 12, // 0x0000000C
            CachedUnlock = 13             // 0x0000000D
        }

        public enum LogonProvider {
            Default,
            WinNT35,
            WinNT40,
            WinNT50
        }

        public enum ImpersonationLevel {
            Anonymous,
            Identify,
            Impersonate,
            Delegate
        }

        public enum BuiltinUser {
            None,
            LocalService,
            NetworkService
        }
    }
}