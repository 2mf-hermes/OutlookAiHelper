using System;
using System.IO;
using System.Text;
using OutlookAiHelper.Adapters.Storage;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Where a cloud source's refresh token is kept, checked against the real thing: the tests
    /// write an actual file under the temporary folder and read it back, because the claim being
    /// pinned is about what ends up on disk.
    ///
    /// The claim that matters is the second one below -- the token is not readable in the file --
    /// and it is checked by looking at the bytes rather than by trusting the call that wrote them.
    /// </summary>
    public static class RefreshTokenStoreTests
    {
        private static string _root;

        public static int Run()
        {
            var failed = 0;
            _root = Path.Combine(Path.GetTempPath(), "oah-refresh-" + Guid.NewGuid().ToString("N"));

            failed += Check("a saved token comes back", () =>
            {
                var store = new RefreshTokenStore(Path.Combine(_root, "basic.dat"));
                store.Save("REFRESH-TOKEN-1");
                return store.Load() == "REFRESH-TOKEN-1" && store.Exists();
            });

            failed += Check("the token is not readable in the file on disk", () =>
            {
                var path = Path.Combine(_root, "sealed.dat");
                new RefreshTokenStore(path).Save("REFRESH-SECRET-0001");

                var bytes = File.ReadAllBytes(path);
                var asText = Encoding.UTF8.GetString(bytes);

                return bytes.Length > 0
                    && asText.IndexOf("REFRESH-SECRET-0001", StringComparison.Ordinal) < 0;
            });

            failed += Check("a store with no file yet reads as no token rather than failing", () =>
            {
                var store = new RefreshTokenStore(Path.Combine(_root, "missing.dat"));
                return !store.Exists() && store.Load() == null;
            });

            failed += Check("a file that is not a token reads as no token rather than failing", () =>
            {
                var junk = Path.Combine(_root, "junk.dat");
                Directory.CreateDirectory(_root);
                File.WriteAllBytes(junk, new byte[64]);
                var junkStore = new RefreshTokenStore(junk);

                var truncated = Path.Combine(_root, "truncated.dat");
                var truncatedStore = new RefreshTokenStore(truncated);
                truncatedStore.Save("REFRESH-TOKEN-2");
                var whole = File.ReadAllBytes(truncated);
                File.WriteAllBytes(truncated, new byte[whole.Length / 2]);

                // The two real cases are the same to the caller: a file copied from another machine
                // or another user cannot be decrypted here, and a damaged one cannot either.
                return junkStore.Load() == null && truncatedStore.Load() == null;
            });

            failed += Check("saving creates the folder it needs", () =>
            {
                var nested = Path.Combine(Path.Combine(_root, "sub", "deeper"), "token.dat");
                var store = new RefreshTokenStore(nested);
                store.Save("REFRESH-TOKEN-3");
                return store.Exists() && store.Load() == "REFRESH-TOKEN-3";
            });

            failed += Check("an empty token is refused instead of stored", () =>
            {
                var store = new RefreshTokenStore(Path.Combine(_root, "empty.dat"));
                return Refused(delegate { store.Save(string.Empty); }) && !store.Exists();
            });

            failed += Check("deleting forgets the token, and doing it twice is harmless", () =>
            {
                var path = Path.Combine(_root, "forget.dat");
                var store = new RefreshTokenStore(path);
                store.Save("REFRESH-TOKEN-4");
                store.Delete();
                var gone = !store.Exists() && store.Load() == null;

                store.Delete();
                store.Save("REFRESH-TOKEN-5");

                return gone && store.Load() == "REFRESH-TOKEN-5";
            });

            failed += Check("a source id that is not a plain name cannot name a token file", () =>
            {
                return RefusedValue(() => JsonPaths.RefreshTokenPath(@"..\..\settings"))
                    && RefusedValue(() => JsonPaths.RefreshTokenPath("a/b"))
                    && RefusedValue(() => JsonPaths.RefreshTokenPath("id.with.dot"))
                    && RefusedValue(() => JsonPaths.RefreshTokenPath("has space"))
                    && RefusedValue(() => JsonPaths.RefreshTokenPath(string.Empty));
            });

            failed += Check("a token file sits under the app's secrets folder and carries the source id", () =>
            {
                const string hex = "0123456789abcdef0123456789abcdef";
                var path = JsonPaths.RefreshTokenPath(hex);
                var localPath = JsonPaths.RefreshTokenPath("local-outlook");

                return path.StartsWith(JsonPaths.SecretsDirectory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    && path.EndsWith(".dat", StringComparison.OrdinalIgnoreCase)
                    && path.IndexOf(hex, StringComparison.Ordinal) > 0
                    && localPath.IndexOf("local-outlook", StringComparison.Ordinal) > 0
                    && !string.Equals(path, localPath, StringComparison.OrdinalIgnoreCase);
            });

            Cleanup();
            return failed;
        }

        private static void Cleanup()
        {
            try
            {
                if (!string.IsNullOrEmpty(_root) && Directory.Exists(_root))
                {
                    Directory.Delete(_root, true);
                }
            }
            catch (Exception)
            {
                // A leftover temporary folder is not a test failure.
            }
        }

        private static bool Refused(Action body)
        {
            try
            {
                body();
                return false;
            }
            catch (ArgumentException)
            {
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool RefusedValue(Func<string> body)
        {
            try
            {
                body();
                return false;
            }
            catch (ArgumentException)
            {
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static int Check(string name, Func<bool> body)
        {
            try
            {
                if (body())
                {
                    Console.WriteLine("PASS " + name);
                    return 0;
                }

                Console.WriteLine("FAIL " + name);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + " :: " + ex.Message);
                return 1;
            }
        }
    }
}
