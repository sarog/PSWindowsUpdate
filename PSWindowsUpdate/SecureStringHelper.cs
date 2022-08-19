using System;
using System.Runtime.InteropServices;
using System.Security;

namespace PSWindowsUpdate {
    [SuppressUnmanagedCodeSecurity]
    internal static class SecureStringHelper {
        internal static unsafe SecureString CreateSecureString(string plainString) {
            if (string.IsNullOrEmpty(plainString)) {
                return new SecureString();
            }

            SecureString secureString;
            fixed (char* chPtr = plainString) {
                secureString = new SecureString(chPtr, plainString.Length);
                secureString.MakeReadOnly();
            }

            return secureString;
        }

        internal static string CreateString(SecureString secureString) {
            if (secureString == null || secureString.Length == 0) {
                return string.Empty;
            }

            var num = IntPtr.Zero;
            string stringBstr;
            try {
                num = Marshal.SecureStringToBSTR(secureString);
                stringBstr = Marshal.PtrToStringBSTR(num);
            } finally {
                if (num != IntPtr.Zero) {
                    Marshal.ZeroFreeBSTR(num);
                }
            }

            return stringBstr;
        }
    }
}