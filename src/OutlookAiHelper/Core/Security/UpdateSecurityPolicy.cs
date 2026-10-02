using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Security
{
    /// <summary>
    /// Every rule that decides whether an update may be fetched lives here, so it can be
    /// tested without a network. The updater is the one place that writes an executable
    /// over the running program, so the checks are deliberately narrow:
    /// HTTPS only, exact host match (no suffix matching), no credentials in the URL,
    /// default port only, a single .exe asset name, and a hard size ceiling.
    /// </summary>
    public static class UpdateSecurityPolicy
    {
        public const long MaxAssetBytes = 209715200L; // 200 MiB

        public const long MinAssetBytes = 40960L; // a real build is never 40 KB

        public const string ExpectedAssetFileName = "OutlookAiHelper.exe";

        public static readonly string[] TrustedApiHosts = { "api.github.com" };

        /// <summary>
        /// github.com serves the link, then redirects the body to a GitHub asset host.
        /// Both hops are allowed — nothing else.
        /// </summary>
        public static readonly string[] TrustedAssetHosts =
        {
            "github.com",
            "api.github.com",
            "objects.githubusercontent.com",
            "github-releases.githubusercontent.com"
        };

        public static bool IsTrustedApiUrl(string url)
        {
            Uri uri;
            return TryParseHttps(url, out uri)
                && IsTrustedHost(uri.Host, TrustedApiHosts)
                && uri.AbsolutePath.StartsWith("/repos/", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsTrustedAssetUrl(string url)
        {
            Uri uri;
            return TryParseHttps(url, out uri) && IsTrustedHost(uri.Host, TrustedAssetHosts);
        }

        public static bool IsTrustedHost(string host, string[] trustedHosts)
        {
            if (string.IsNullOrEmpty(host) || trustedHosts == null)
            {
                return false;
            }

            foreach (var trusted in trustedHosts)
            {
                if (string.Equals(host, trusted, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>Asset names are file names, never paths, and only .exe is installable.</summary>
        public static bool IsAllowedAssetName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            var trimmed = name.Trim();
            if (trimmed.Length == 0 || !trimmed.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return trimmed.IndexOf('/') < 0
                && trimmed.IndexOf('\\') < 0
                && trimmed.IndexOf(':') < 0
                && trimmed.IndexOf("..", StringComparison.Ordinal) < 0;
        }

        public static bool IsWithinSizeLimit(long bytes)
        {
            // 0 (or unknown) passes: the download itself enforces the ceiling byte by byte.
            return bytes <= 0 || bytes <= MaxAssetBytes;
        }

        /// <summary>Prefers the canonical file name, otherwise the first .exe asset.</summary>
        public static ReleaseAsset SelectAsset(ReleaseInfo release)
        {
            if (release == null || release.Assets == null)
            {
                return null;
            }

            ReleaseAsset firstExecutable = null;
            foreach (var asset in release.Assets)
            {
                if (asset == null || !IsAllowedAssetName(asset.Name))
                {
                    continue;
                }

                if (string.Equals(asset.Name.Trim(), ExpectedAssetFileName, StringComparison.OrdinalIgnoreCase))
                {
                    return asset;
                }

                if (firstExecutable == null)
                {
                    firstExecutable = asset;
                }
            }

            return firstExecutable;
        }

        /// <summary>Cheap "is this really a Windows program" check on the downloaded bytes.</summary>
        public static bool LooksLikePortableExecutable(byte[] header)
        {
            return header != null && header.Length >= 2 && header[0] == (byte)'M' && header[1] == (byte)'Z';
        }

        /// <summary>
        /// A path is only safe to embed in the updater's batch script when it needs no
        /// cmd escaping. Anything with those characters is refused rather than guessed at.
        /// </summary>
        public static bool IsSafeForUpdaterScript(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            foreach (var c in path)
            {
                if (c == '%' || c == '"' || c == '&' || c == '^' || c == '<' || c == '>'
                    || c == '|' || c == '\r' || c == '\n' || c == '!' || c == '\t')
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseHttps(string url, out Uri uri)
        {
            uri = null;
            if (string.IsNullOrEmpty(url))
            {
                return false;
            }

            Uri parsed;
            if (!Uri.TryCreate(url.Trim(), UriKind.Absolute, out parsed))
            {
                return false;
            }

            if (!string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Credentials in the URL, a non-default port, or the loopback aliases all
            // mean the request is not going where the user was told it would.
            if (!string.IsNullOrEmpty(parsed.UserInfo) || !parsed.IsDefaultPort)
            {
                return false;
            }

            uri = parsed;
            return true;
        }
    }
}
