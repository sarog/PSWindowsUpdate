using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace PSWindowsUpdate {
    public class NativeMethods {
        [DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool CredRead(
            string target,
            CredentialType type,
            int reservedFlag,
            out IntPtr credentialPtr);

        [DllImport("Advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool CredWrite([In] ref CREDENTIAL userCredential, [In] uint flags);

        [DllImport("Advapi32.dll", SetLastError = true)]
        internal static extern void CredFree([In] IntPtr cred);

        [DllImport("Advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode)]
        internal static extern bool CredDelete(StringBuilder target, CredentialType type, int flags);

        [DllImport("Advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool CredEnumerate(
            string filter,
            int flag,
            out int count,
            out IntPtr pCredentials);

        internal static IEnumerable<CREDENTIAL> CredEnumerate() {
            int count;
            IntPtr pCredentials;
            if (!CredEnumerate((string)null, 0, out count, out pCredentials)) {
                throw new Exception("Failed to enumerate credentials");
            }

            var source = new IntPtr[count];
            for (var index = 0; index < count; ++index) {
                source[index] = Marshal.ReadIntPtr(pCredentials, index * Marshal.SizeOf(typeof(IntPtr)));
            }

            return ((IEnumerable<IntPtr>)source).Select<IntPtr, CREDENTIAL>((Func<IntPtr, CREDENTIAL>)(ptr => (CREDENTIAL)Marshal.PtrToStructure(ptr, typeof(CREDENTIAL))));
        }

        internal struct CREDENTIAL {
            public int Flags;
            public int Type;
            [MarshalAs(UnmanagedType.LPWStr)] public string TargetName;
            [MarshalAs(UnmanagedType.LPWStr)] public string Comment;
            public long LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            [MarshalAs(UnmanagedType.LPWStr)] public string TargetAlias;
            [MarshalAs(UnmanagedType.LPWStr)] public string UserName;
        }

        internal sealed class CriticalCredentialHandle : CriticalHandleZeroOrMinusOneIsInvalid {
            internal CriticalCredentialHandle(IntPtr preexistingHandle) {
                SetHandle(preexistingHandle);
            }

            internal CREDENTIAL GetCredential() {
                if (!IsInvalid) {
                    return (CREDENTIAL)Marshal.PtrToStructure(handle, typeof(CREDENTIAL));
                }

                throw new InvalidOperationException("Invalid CriticalHandle!");
            }

            protected override bool ReleaseHandle() {
                if (IsInvalid) {
                    return false;
                }

                CredFree(handle);
                SetHandleAsInvalid();
                return true;
            }
        }
    }
}