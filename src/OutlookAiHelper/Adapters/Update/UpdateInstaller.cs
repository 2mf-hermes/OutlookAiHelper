using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Core.Security;

namespace OutlookAiHelper.Adapters.Update
{
    /// <summary>
    /// Downloads an approved release asset and hands the swap over to a tiny batch script,
    /// because a running .exe cannot overwrite itself. Nothing here runs unless the user
    /// pressed the install button: the class only ever touches a verified URL, a temp file
    /// and the folder the app already runs from.
    /// </summary>
    public sealed class UpdateInstaller
    {
        private const int TimeoutMs = 60000;
        private const int BufferSize = 81920;

        /// <summary>
        /// Downloads the asset to %TEMP%. Returns the file path, or null with an error key.
        /// Progress is reported as 0..100 (never called when the server sends no length).
        /// </summary>
        public string Download(string assetUrl, string assetName, long expectedBytes, Action<int> onPercent, out string errorKey)
        {
            errorKey = null;

            if (!UpdateSecurityPolicy.IsTrustedAssetUrl(assetUrl)
                || !UpdateSecurityPolicy.IsAllowedAssetName(assetName))
            {
                errorKey = "update.error.untrusted";
                return null;
            }

            var target = Path.Combine(
                Path.GetTempPath(),
                "OutlookAiHelper-" + Guid.NewGuid().ToString("N") + ".exe");

            try
            {
                // Must happen here as well: the download may be the first request this
                // process ever makes, and a fresh process here defaults to Ssl3|Tls,
                // which the asset host refuses.
                UpdateHttp.EnsureModernTls();

                var request = (HttpWebRequest)WebRequest.Create(assetUrl);
                request.Method = "GET";
                request.UserAgent = "OutlookAiHelper/" + AppInfo.Version;
                request.Timeout = TimeoutMs;
                request.ReadWriteTimeout = TimeoutMs;
                request.AllowAutoRedirect = true; // github.com -> objects.githubusercontent.com

                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    // The redirect target must clear the same policy as the original link.
                    var finalUrl = response.ResponseUri == null ? assetUrl : response.ResponseUri.AbsoluteUri;
                    if (!UpdateSecurityPolicy.IsTrustedAssetUrl(finalUrl))
                    {
                        errorKey = "update.error.untrusted";
                        return null;
                    }

                    var total = response.ContentLength;
                    if (total > UpdateSecurityPolicy.MaxAssetBytes)
                    {
                        errorKey = "update.error.tooLarge";
                        return null;
                    }

                    long written = 0;
                    var buffer = new byte[BufferSize];
                    using (var input = response.GetResponseStream())
                    using (var output = File.Create(target))
                    {
                        int read;
                        while (input != null && (read = input.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            written += read;
                            if (written > UpdateSecurityPolicy.MaxAssetBytes)
                            {
                                output.Close();
                                TryDelete(target);
                                errorKey = "update.error.tooLarge";
                                return null;
                            }

                            output.Write(buffer, 0, read);
                            if (onPercent != null && total > 0)
                            {
                                var percent = (int)(written * 100 / total);
                                onPercent(percent > 100 ? 100 : percent);
                            }
                        }
                    }
                }

                var size = new FileInfo(target).Length;
                if (!VerifyDownload(target, size, expectedBytes, out errorKey))
                {
                    TryDelete(target);
                    return null;
                }

                FileLogger.Info("Update downloaded: " + size.ToString(CultureInfo.InvariantCulture) + " bytes");
                return target;
            }
            catch (WebException ex)
            {
                FileLogger.Error("UpdateDownload", ex);
                errorKey = ex.Status == WebExceptionStatus.Timeout ? "update.error.timeout" : "update.error.network";
            }
            catch (Exception ex)
            {
                FileLogger.Error("UpdateDownload", ex);
                errorKey = "update.error.download";
            }

            TryDelete(target);
            return null;
        }

        /// <summary>
        /// Rejects truncated downloads and files that are not Windows executables. The
        /// GitHub API publishes no checksum, so size + MZ header are the honest checks
        /// available; anything else would be theatre.
        /// </summary>
        public static bool VerifyDownload(string path, long size, long expectedBytes, out string errorKey)
        {
            errorKey = null;

            if (size < UpdateSecurityPolicy.MinAssetBytes || size > UpdateSecurityPolicy.MaxAssetBytes)
            {
                errorKey = "update.error.incomplete";
                return false;
            }

            if (expectedBytes > 0 && size != expectedBytes)
            {
                errorKey = "update.error.size";
                return false;
            }

            var header = new byte[2];
            using (var stream = File.OpenRead(path))
            {
                if (stream.Read(header, 0, 2) != 2 || !UpdateSecurityPolicy.LooksLikePortableExecutable(header))
                {
                    errorKey = "update.error.incomplete";
                    return false;
                }
            }

            return true;
        }

        /// <summary>The folder holding the running .exe has to be writable to self-update.</summary>
        public static bool CanWriteToApplicationFolder(out string errorKey)
        {
            errorKey = null;
            try
            {
                var exe = CurrentExecutablePath();
                if (string.IsNullOrEmpty(exe))
                {
                    errorKey = "update.error.locate";
                    return false;
                }

                var folder = Path.GetDirectoryName(exe);
                if (string.IsNullOrEmpty(folder))
                {
                    errorKey = "update.error.locate";
                    return false;
                }

                var probe = Path.Combine(folder, ".oah-write-test-" + Guid.NewGuid().ToString("N"));
                using (var stream = File.Create(probe))
                {
                    stream.WriteByte(0);
                }

                File.Delete(probe);
                return true;
            }
            catch (Exception)
            {
                errorKey = "update.error.readonly";
                return false;
            }
        }

        /// <summary>
        /// Starts a detached script that waits for this process to exit, replaces the
        /// binary, relaunches it and removes its own leftovers.
        /// </summary>
        public bool TryStartSwap(string downloadedExe, out string errorKey)
        {
            errorKey = null;

            var currentExe = CurrentExecutablePath();
            if (string.IsNullOrEmpty(currentExe))
            {
                errorKey = "update.error.locate";
                return false;
            }

            if (string.IsNullOrEmpty(downloadedExe) || !File.Exists(downloadedExe))
            {
                errorKey = "update.error.download";
                return false;
            }

            if (!UpdateSecurityPolicy.IsSafeForUpdaterScript(currentExe)
                || !UpdateSecurityPolicy.IsSafeForUpdaterScript(downloadedExe))
            {
                errorKey = "update.error.unsupportedPath";
                return false;
            }

            try
            {
                var script = Path.Combine(
                    Path.GetTempPath(),
                    "OutlookAiHelper-update-" + Guid.NewGuid().ToString("N") + ".cmd");
                File.WriteAllText(script, BuildSwapScript(currentExe, downloadedExe), Encoding.Default);

                var cmd = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "cmd.exe");

                var startInfo = new ProcessStartInfo
                {
                    FileName = File.Exists(cmd) ? cmd : "cmd.exe",
                    Arguments = "/c \"" + script + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = Path.GetTempPath()
                };

                Process.Start(startInfo);
                FileLogger.Info("Update swap started for " + downloadedExe);
                return true;
            }
            catch (Exception ex)
            {
                FileLogger.Error("UpdateSwap", ex);
                errorKey = "update.error.download";
                return false;
            }
        }

        /// <summary>Folder the running program lives in, used as the restart working directory.</summary>
        private static string FolderOf(string exePath)
        {
            try
            {
                return Path.GetDirectoryName(exePath);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static string BuildSwapScript(string currentExe, string downloadedExe)
        {
            var builder = new StringBuilder();
            builder.AppendLine("@echo off");
            builder.AppendLine("setlocal");
            builder.AppendLine("set /a tries=0");
            builder.AppendLine(":retry");
            builder.AppendLine("copy /y \"" + downloadedExe + "\" \"" + currentExe + "\" >nul 2>&1");
            builder.AppendLine("if not errorlevel 1 goto installed");
            builder.AppendLine("set /a tries+=1");
            builder.AppendLine("if %tries% geq 90 goto failed");
            builder.AppendLine("ping -n 2 127.0.0.1 >nul");
            builder.AppendLine("goto retry");
            builder.AppendLine(":installed");
            var home = FolderOf(currentExe);
            if (!string.IsNullOrEmpty(home))
            {
                builder.AppendLine("cd /d \"" + home + "\"");
            }

            builder.AppendLine("start \"\" \"" + currentExe + "\"");
            builder.AppendLine("del \"" + downloadedExe + "\" >nul 2>&1");
            builder.AppendLine("del \"%~f0\" >nul 2>&1");
            builder.AppendLine("exit /b 0");
            builder.AppendLine(":failed");
            builder.AppendLine("exit /b 1");
            return builder.ToString();
        }

        public static string CurrentExecutablePath()
        {
            try
            {
                var module = Process.GetCurrentProcess().MainModule;
                return module == null ? null : module.FileName;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
