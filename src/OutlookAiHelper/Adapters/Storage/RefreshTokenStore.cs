using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAiHelper.Adapters.Storage
{
    /// <summary>
    /// Holds one cloud source's refresh token in its own file, encrypted with the current Windows
    /// user's key (DPAPI, CurrentUser scope).
    ///
    /// The refresh token is the one real credential this app ever holds: it is what lets a mailbox
    /// be read again tomorrow without the user signing in again. So it is kept out of settings.json
    /// -- a plain json file that gets copied between machines, pasted into bug reports and restored
    /// from backups -- and out of the log, and it is encrypted at rest. DPAPI CurrentUser means only
    /// this Windows account on this machine can turn the file back into a token: a copy of the file
    /// elsewhere is unreadable, and the way to read the same mailbox on another machine is to sign
    /// in there, which the device code flow exists to make cheap.
    ///
    /// No token ever appears in a message this class can produce. The paths and the reason are
    /// enough to describe a failure without echoing a secret.
    /// </summary>
    internal sealed class RefreshTokenStore
    {
        private readonly string _path;

        /// <summary>
        /// The store for one file. The path is passed in rather than derived here so this class
        /// carries the crypto and nothing else; where the file belongs is the layout's business
        /// (JsonPaths.RefreshTokenPath), and a test can point this at a temporary file.
        /// </summary>
        internal RefreshTokenStore(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("a file path is required", "path");
            }

            _path = path;
        }

        internal string FilePath
        {
            get { return _path; }
        }

        internal bool Exists()
        {
            return File.Exists(_path);
        }

        /// <summary>Writes the token, replacing whatever was there. Creates the folder it needs.</summary>
        internal void Save(string refreshToken)
        {
            if (string.IsNullOrEmpty(refreshToken))
            {
                // Storing an empty string would look exactly like a successful sign-in that is not
                // one, and the failure would surface later as "the mailbox stopped working".
                throw new ArgumentException("a token is required", "refreshToken");
            }

            var folder = Path.GetDirectoryName(_path);
            if (!string.IsNullOrEmpty(folder))
            {
                Directory.CreateDirectory(folder);
            }

            File.WriteAllBytes(_path, ProtectedData.Protect(
                Encoding.UTF8.GetBytes(refreshToken), null, DataProtectionScope.CurrentUser));
        }

        /// <summary>
        /// The stored token, or null when there is none or it cannot be read. Unreadable covers the
        /// real cases -- the file was copied from another machine or another user, or it was
        /// damaged -- and all of them mean the same thing to the caller: sign in again. None of them
        /// is worth an exception on the UI thread.
        /// </summary>
        internal string Load()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    return null;
                }

                var token = Encoding.UTF8.GetString(ProtectedData.Unprotect(
                    File.ReadAllBytes(_path), null, DataProtectionScope.CurrentUser));

                return string.IsNullOrEmpty(token) ? null : token;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Forgets the token: used when a source is removed, and when the service refuses the
        /// sign-in and the stored token is what has to stop being tried.
        /// </summary>
        internal void Delete()
        {
            try
            {
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
            }
            catch (Exception)
            {
                // A token that cannot be deleted is still a token that will not be used, because the
                // caller drops the source. Failing here would turn housekeeping into a crash.
            }
        }
    }
}
