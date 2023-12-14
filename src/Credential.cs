using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;

namespace PSWindowsUpdate {
    /// <summary>
    ///     Class Credential, wrapper for native CREDENTIAL structure.
    ///     See CREDENTIAL structure
    ///     <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/aa374788(v=vs.85).aspx">documentation.</see>
    ///     See Credential Manager
    ///     <see href="http://windows.microsoft.com/en-us/windows7/what-is-credential-manager">documentation.</see>
    /// </summary>
    public class Credential : IDisposable {
        /// <summary>The lock object</summary>
        private static readonly object LockObject = new object();

        /// <summary>The unmanaged code permission</summary>
        private static readonly SecurityPermission UnmanagedCodePermission;

        /// <summary>The credential description</summary>
        private string description;

        /// <summary>The disposed flag</summary>
        private bool disposed;

        /// <summary>The last write time</summary>
        private DateTime lastWriteTime;

        /// <summary>The password</summary>
        private SecureString password;

        /// <summary>The persistence type</summary>
        private PersistenceType persistenceType;

        /// <summary>
        ///     The string that contains the name of the credential
        /// </summary>
        private string target;

        /// <summary>The credential type</summary>
        private CredentialType type;

        /// <summary>The username</summary>
        private string username;

        /// <summary>
        ///     Initializes UnmanagedCodePermission for the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        static Credential() {
            lock (LockObject) {
                UnmanagedCodePermission = new SecurityPermission(SecurityPermissionFlag.UnmanagedCode);
            }
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        public Credential()
            : this(null) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        /// <param name="username">The username.</param>
        public Credential(string username)
            : this(username, null) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        public Credential(string username, string password)
            : this(username, password, null) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        /// <param name="target">The string that contains the name of the credential.</param>
        public Credential(string username, string password, string target)
            : this(username, password, target, CredentialType.Generic) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        /// <param name="target">The string that contains the name of the credential.</param>
        /// <param name="type">The credential type.</param>
        public Credential(string username, string password, string target, CredentialType type) {
            Username = username;
            Password = password;
            Target = target;
            Type = type;
            PersistenceType = PersistenceType.Session;
            lastWriteTime = DateTime.MinValue;
        }

        /// <summary>Gets or sets the username.</summary>
        /// <value>The user name of the account used to connect to TargetName.</value>
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

        /// <summary>Gets or sets the password.</summary>
        /// <value>The decoded secure string password.</value>
        public string Password {
            get => SecureStringHelper.CreateString(SecurePassword);
            set {
                CheckNotDisposed();
                SecurePassword = SecureStringHelper.CreateSecureString(string.IsNullOrEmpty(value) ? string.Empty : value);
            }
        }

        /// <summary>Gets or sets the secure password.</summary>
        /// <value>The secure password of the account used to connect to TargetName.</value>
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

        /// <summary>Gets or sets the target.</summary>
        /// <value>
        ///     The name of the credential. The TargetName and Type members uniquely identify the credential. This member cannot
        ///     be changed after the credential is created. Instead, the credential with the old name should be deleted and the
        ///     credential with the new name created.
        /// </value>
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

        /// <summary>Gets or sets the description.</summary>
        /// <value>
        ///     The string comment from the user that describes this credential. This member cannot be longer than
        ///     CRED_MAX_STRING_LENGTH (256) characters.
        /// </value>
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

        /// <summary>Gets the last write time.</summary>
        /// <value>The last write time.</value>
        public DateTime LastWriteTime => LastWriteTimeUtc.ToLocalTime();

        /// <summary>Gets the last write time UTC.</summary>
        /// <value>The last write time UTC.</value>
        public DateTime LastWriteTimeUtc {
            get {
                CheckNotDisposed();
                return lastWriteTime;
            }
            private set => lastWriteTime = value;
        }

        /// <summary>Gets or sets the type.</summary>
        /// <value>The type of the credential. This member cannot be changed after the credential is created.</value>
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

        /// <summary>Gets or sets the type of the persistence.</summary>
        /// <value>Defines the persistence of this credential. This member can be read and written.</value>
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

        /// <summary>
        ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Finalizes an instance of the <see cref="T:PSWindowsUpdate.Credential" /> class.
        /// </summary>
        ~Credential() {
            Dispose(false);
        }

        /// <summary>
        ///     Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing">
        ///     <c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only
        ///     unmanaged resources.
        /// </param>
        private void Dispose(bool disposing) {
            if (!disposed && disposing) {
                SecurePassword.Clear();
                SecurePassword.Dispose();
            }

            disposed = true;
        }

        /// <summary>Ensures this instance is not disposed.</summary>
        /// <exception cref="T:System.ObjectDisposedException">Credential object is already disposed.</exception>
        private void CheckNotDisposed() {
            if (disposed) {
                throw new ObjectDisposedException("Credential object is already disposed.");
            }
        }

        /// <summary>Saves this instance.</summary>
        /// <returns><c>true</c> if credential is saved properly, <c>false</c> otherwise.</returns>
        /// <exception cref="T:System.ArgumentOutOfRangeException">password;The password has exceeded 512 bytes.</exception>
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

        /// <summary>Deletes this instance.</summary>
        /// <returns><c>true</c> if credential was deleted properly, <c>false</c> otherwise.</returns>
        /// <exception cref="T:System.InvalidOperationException">Target must be specified to delete a credential.</exception>
        public bool Delete() {
            CheckNotDisposed();
            UnmanagedCodePermission.Demand();
            return !string.IsNullOrEmpty(Target)
                ? NativeMethods.CredDelete(string.IsNullOrEmpty(Target) ? new StringBuilder() : new StringBuilder(Target), Type, 0)
                : throw new InvalidOperationException("Target must be specified to delete a credential.");
        }

        /// <summary>Loads this instance.</summary>
        /// <returns><c>true</c> if credential is load properly, <c>false</c> otherwise.</returns>
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

        /// <summary>
        ///     Checks if the credential with these properties exists
        /// </summary>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        /// <exception cref="T:System.InvalidOperationException">Target must be specified to check existance of a credential.</exception>
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

        /// <summary>Loads all credentials</summary>
        public static IEnumerable<Credential> LoadAll() {
            UnmanagedCodePermission.Demand();
            return NativeMethods.CredEnumerate()
                .Select(c => new Credential(c.UserName, null, c.TargetName))
                .Where(c => c.Load());
        }

        /// <summary>Loads the internal.</summary>
        /// <param name="credential">The credential.</param>
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

        /// <summary>Loads the internal.</summary>
        public override string ToString() {
            return string.Format("Username: {0}, Target: {1}, LastWriteTime: {2}, LastWriteTimeUtc: {3}, Type: {4}, PersistenceType: {5}", Username, Target,
                LastWriteTime, LastWriteTimeUtc, Type, PersistenceType);
        }
    }
}