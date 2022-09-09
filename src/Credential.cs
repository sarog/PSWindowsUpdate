using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;

namespace PSWindowsUpdate {
    public class Credential : IDisposable {
        private static readonly object LockObject = new object();
        private static readonly SecurityPermission UnmanagedCodePermission;
        private string description;
        private bool disposed;
        private DateTime lastWriteTime;
        private SecureString password;
        private PersistenceType persistenceType;
        private string target;
        private CredentialType type;
        private string username;

        static Credential() {
            lock (LockObject) {
                UnmanagedCodePermission = new SecurityPermission(SecurityPermissionFlag.UnmanagedCode);
            }
        }

        public Credential()
            : this(null) { }

        public Credential(string username)
            : this(username, null) { }

        public Credential(string username, string password)
            : this(username, password, null) { }

        public Credential(string username, string password, string target)
            : this(username, password, target, CredentialType.Generic) { }

        public Credential(string username, string password, string target, CredentialType type) {
            Username = username;
            Password = password;
            Target = target;
            Type = type;
            PersistenceType = PersistenceType.Session;
            lastWriteTime = DateTime.MinValue;
        }

        public string Username {
            get {
                CheckNotDisposed();
                return username;
            }
            set {
                CheckNotDisposed();
                username = value;
            }
        }

        public string Password {
            get => SecureStringHelper.CreateString(SecurePassword);
            set {
                CheckNotDisposed();
                SecurePassword = SecureStringHelper.CreateSecureString(string.IsNullOrEmpty(value) ? string.Empty : value);
            }
        }

        public SecureString SecurePassword {
            get {
                CheckNotDisposed();
                UnmanagedCodePermission.Demand();
                return password == null ? new SecureString() : password.Copy();
            }
            set {
                CheckNotDisposed();
                if (password != null) {
                    password.Clear();
                    password.Dispose();
                }

                password = value == null ? new SecureString() : value.Copy();
            }
        }

        public string Target {
            get {
                CheckNotDisposed();
                return target;
            }
            set {
                CheckNotDisposed();
                target = value;
            }
        }

        public string Description {
            get {
                CheckNotDisposed();
                return description;
            }
            set {
                CheckNotDisposed();
                description = value;
            }
        }

        public DateTime LastWriteTime => LastWriteTimeUtc.ToLocalTime();

        public DateTime LastWriteTimeUtc {
            get {
                CheckNotDisposed();
                return lastWriteTime;
            }
            private set => lastWriteTime = value;
        }

        public CredentialType Type {
            get {
                CheckNotDisposed();
                return type;
            }
            set {
                CheckNotDisposed();
                type = value;
            }
        }

        public PersistenceType PersistenceType {
            get {
                CheckNotDisposed();
                return persistenceType;
            }
            set {
                CheckNotDisposed();
                persistenceType = value;
            }
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Credential() {
            Dispose(false);
        }

        private void Dispose(bool disposing) {
            if (!disposed && disposing) {
                SecurePassword.Clear();
                SecurePassword.Dispose();
            }

            disposed = true;
        }

        private void CheckNotDisposed() {
            if (disposed) {
                throw new ObjectDisposedException("Credential object is already disposed.");
            }
        }

        public bool Save() {
            CheckNotDisposed();
            UnmanagedCodePermission.Demand();
            var bytes = Encoding.Unicode.GetBytes(Password);
            if (Password.Length > 512) {
                throw new ArgumentOutOfRangeException("password", "The password has exceeded 512 bytes.");
            }

            var userCredential = new NativeMethods.CREDENTIAL() {
                TargetName = Target,
                UserName = Username,
                CredentialBlob = Marshal.StringToCoTaskMemUni(Password),
                CredentialBlobSize = bytes.Length,
                Comment = Description,
                Type = (int)Type,
                Persist = (int)PersistenceType
            };
            
            if (!NativeMethods.CredWrite(ref userCredential, 0U)) {
                return false;
            }

            LastWriteTimeUtc = DateTime.UtcNow;
            return true;
        }

        public bool Delete() {
            CheckNotDisposed();
            UnmanagedCodePermission.Demand();
            return !string.IsNullOrEmpty(Target)
                ? NativeMethods.CredDelete(string.IsNullOrEmpty(Target) ? new StringBuilder() : new StringBuilder(Target), Type, 0)
                : throw new InvalidOperationException("Target must be specified to delete a credential.");
        }

        public bool Load() {
            CheckNotDisposed();
            UnmanagedCodePermission.Demand();
            IntPtr credentialPtr;
            if (!NativeMethods.CredRead(Target, Type, 0, out credentialPtr)) {
                return false;
            }

            using (var credentialHandle = new NativeMethods.CriticalCredentialHandle(credentialPtr)) {
                LoadInternal(credentialHandle.GetCredential());
            }

            return true;
        }

        public bool Exists() {
            CheckNotDisposed();
            UnmanagedCodePermission.Demand();
            if (string.IsNullOrEmpty(Target)) {
                throw new InvalidOperationException("Target must be specified to check existance of a credential.");
            }

            using (var credential = new Credential() {
                       Target = Target,
                       Type = Type
                   }) {
                return credential.Load();
            }
        }

        public static IEnumerable<Credential> LoadAll() {
            UnmanagedCodePermission.Demand();
            return NativeMethods.CredEnumerate()
                .Select<NativeMethods.CREDENTIAL, Credential>((Func<NativeMethods.CREDENTIAL, Credential>)(c => new Credential(c.UserName, (string)null, c.TargetName)))
                .Where<Credential>((Func<Credential, bool>)(c => c.Load()));
        }

        internal void LoadInternal(NativeMethods.CREDENTIAL credential) {
            Username = credential.UserName;
            if (credential.CredentialBlobSize > 0) {
                Password = Marshal.PtrToStringUni(credential.CredentialBlob, credential.CredentialBlobSize / 2);
            }

            Target = credential.TargetName;
            Type = (CredentialType)credential.Type;
            PersistenceType = (PersistenceType)credential.Persist;
            Description = credential.Comment;
            LastWriteTimeUtc = DateTime.FromFileTimeUtc(credential.LastWritten);
        }

        public override string ToString() {
            return string.Format("Username: {0}, Target: {1}, LastWriteTime: {2}, LastWriteTimeUtc: {3}, Type: {4}, PersistenceType: {5}", Username, Target,
                LastWriteTime, LastWriteTimeUtc, Type, PersistenceType);
        }
    }
}