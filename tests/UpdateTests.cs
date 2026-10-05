using System;
using System.Collections.Generic;
using OutlookAiHelper.Adapters.Update;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Core.Security;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The updater is the only part of the app that downloads and then runs new code, so
    /// its decisions are pinned here: version comparison, download-source allowlist, asset
    /// selection, payload shape and the "may this be auto-installed?" hand-off to the UI.
    /// Nothing in this file touches the network.
    /// </summary>
    public static class UpdateTests
    {
        public static int Run()
        {
            var failed = 0;
            failed += VersionComparison();
            failed += SourcePolicy();
            failed += AssetSelection();
            failed += PayloadChecks();
            failed += FeedParsing();
            failed += OfferDecision();
            failed += FeedFailureHandling();
            return failed;
        }

        private static int VersionComparison()
        {
            var failed = 0;

            failed += Check("tags may carry a v and omit trailing fields", () =>
            {
                var shortForm = AppVersion.ParseOrNull("v1.2");
                var longForm = AppVersion.ParseOrNull("1.2.0");
                return shortForm != null
                    && longForm != null
                    && shortForm.CompareTo(longForm) == 0
                    && shortForm.Normalized == "1.2.0";
            });

            failed += Check("a newer release wins and an equal one does not", () =>
                AppVersion.ParseOrNull("1.0.1").IsNewerThan(AppVersion.ParseOrNull("1.0.0"))
                && AppVersion.ParseOrNull("2.0.0").IsNewerThan(AppVersion.ParseOrNull("1.9.9"))
                && AppVersion.ParseOrNull("1.0.0.1").IsNewerThan(AppVersion.ParseOrNull("1.0.0"))
                && !AppVersion.ParseOrNull("1.0.0").IsNewerThan(AppVersion.ParseOrNull("1.0.0"))
                && !AppVersion.ParseOrNull("1.0.0").IsNewerThan(AppVersion.ParseOrNull("1.0.1")));

            failed += Check("a release outranks its own prereleases", () =>
                AppVersion.ParseOrNull("1.0.0").IsNewerThan(AppVersion.ParseOrNull("1.0.0-rc.2"))
                && AppVersion.ParseOrNull("1.0.0-rc.2").IsNewerThan(AppVersion.ParseOrNull("1.0.0-rc.1"))
                && !AppVersion.ParseOrNull("1.0.0-rc.1").IsNewerThan(AppVersion.ParseOrNull("1.0.0")));

            failed += Check("build metadata is ignored and junk is refused", () =>
                AppVersion.ParseOrNull("1.0.0+build.7").Normalized == "1.0.0"
                && AppVersion.ParseOrNull(" 1.0.1 ").Normalized == "1.0.1"
                && AppVersion.ParseOrNull("abc") == null
                && AppVersion.ParseOrNull("1.x.0") == null
                && AppVersion.ParseOrNull("1.0.0.0.0") == null
                && AppVersion.ParseOrNull(string.Empty) == null
                && AppVersion.ParseOrNull(null) == null);

            return failed;
        }

        private static int SourcePolicy()
        {
            var failed = 0;

            failed += Check("only GitHub hosts may serve the download", () =>
                UpdateSecurityPolicy.IsTrustedAssetUrl(
                    "https://github.com/2mf-hermes/OutlookAiHelper/releases/download/v1.0.1/OutlookAiHelper.exe")
                && UpdateSecurityPolicy.IsTrustedAssetUrl("https://objects.githubusercontent.com/github-production-release-asset/1/2")
                // The host github.com actually redirects to today (observed on the v1.1.0
                // download). Without it the app refuses its own honest release link.
                && UpdateSecurityPolicy.IsTrustedAssetUrl(
                    "https://release-assets.githubusercontent.com/github-production-release-asset/1/2?sp=r&sig=x")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl(
                    "https://release-assets.githubusercontent.com.evil.example/github-production-release-asset/1/2")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl(
                    "https://evil.example/release-assets.githubusercontent.com/x.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl(
                    "http://github.com/2mf-hermes/OutlookAiHelper/releases/download/v1.0.1/OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl("https://github.com.evil.example/OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl("https://raw.githubusercontent.com/x/OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl("https://evil.example/github.com/OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl(null));

            failed += Check("credentials, odd ports and non-repo API paths are refused", () =>
                !UpdateSecurityPolicy.IsTrustedAssetUrl("https://user:pass@github.com/OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsTrustedAssetUrl("https://github.com:8443/OutlookAiHelper.exe")
                && UpdateSecurityPolicy.IsTrustedApiUrl("https://api.github.com/repos/2mf-hermes/OutlookAiHelper/releases/latest")
                && !UpdateSecurityPolicy.IsTrustedApiUrl("https://api.github.com/user")
                && !UpdateSecurityPolicy.IsTrustedApiUrl("https://api.github.com.evil.example/repos/a/b/releases/latest"));

            failed += Check("both updater paths opt into modern TLS", () =>
            {
                // Reproduces the real failure this pins: a fresh process here defaults to
                // Ssl3|Tls, which the GitHub API and the asset host both refuse. The check
                // and the download used to rely on ordering between them; now one shared
                // helper guarantees it. Idempotent, so calling it here is safe either way.
                UpdateHttp.EnsureModernTls();
                return (System.Net.ServicePointManager.SecurityProtocol
                    & (System.Net.SecurityProtocolType)3072) == (System.Net.SecurityProtocolType)3072;
            });

            failed += Check("only a bare .exe file name may be installed", () =>
                UpdateSecurityPolicy.IsAllowedAssetName("OutlookAiHelper.exe")
                && UpdateSecurityPolicy.IsAllowedAssetName("OutlookAiHelper v1.0.1.exe")
                && !UpdateSecurityPolicy.IsAllowedAssetName("OutlookAiHelper.exe.bat")
                && !UpdateSecurityPolicy.IsAllowedAssetName("..\\OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsAllowedAssetName("C:\\temp\\OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsAllowedAssetName("notes.txt")
                && !UpdateSecurityPolicy.IsAllowedAssetName(string.Empty));

            failed += Check("paths embedded in the swap script need no cmd escaping", () =>
                UpdateSecurityPolicy.IsSafeForUpdaterScript("C:\\Users\\MINFU LIAO\\Desktop\\OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsSafeForUpdaterScript("C:\\temp\\100%\\OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsSafeForUpdaterScript("C:\\temp\\a\"b\\OutlookAiHelper.exe")
                && !UpdateSecurityPolicy.IsSafeForUpdaterScript(null));

            return failed;
        }

        private static int AssetSelection()
        {
            var failed = 0;

            failed += Check("the canonical exe wins over other assets", () =>
            {
                var release = Release(
                    "v1.0.1",
                    "https://github.com/2mf-hermes/OutlookAiHelper/releases/download/v1.0.1/OutlookAiHelper.exe",
                    5242880);
                release.Assets.Insert(0, new ReleaseAsset { Name = "checksums.txt" });
                release.Assets.Add(new ReleaseAsset { Name = "other.exe", DownloadUrl = "https://github.com/y", Size = 2000 });

                var selected = UpdateSecurityPolicy.SelectAsset(release);
                return selected != null && selected.Name == "OutlookAiHelper.exe";
            });

            failed += Check("nothing installable is selected when no executable is attached", () =>
                UpdateSecurityPolicy.SelectAsset(new ReleaseInfo()) == null
                && UpdateSecurityPolicy.SelectAsset(null) == null);

            return failed;
        }

        private static int PayloadChecks()
        {
            var failed = 0;

            failed += Check("downloaded bytes have to look like a Windows program", () =>
                UpdateSecurityPolicy.LooksLikePortableExecutable(new byte[] { (byte)'M', (byte)'Z' })
                && !UpdateSecurityPolicy.LooksLikePortableExecutable(new byte[] { (byte)'P', (byte)'K' })
                && !UpdateSecurityPolicy.LooksLikePortableExecutable(new byte[] { (byte)'M' })
                && !UpdateSecurityPolicy.LooksLikePortableExecutable(null));

            failed += Check("the size ceiling refuses oversized assets", () =>
                UpdateSecurityPolicy.IsWithinSizeLimit(0)
                && UpdateSecurityPolicy.IsWithinSizeLimit(5L * 1024 * 1024)
                && !UpdateSecurityPolicy.IsWithinSizeLimit(UpdateSecurityPolicy.MaxAssetBytes + 1));

            return failed;
        }

        private static int FeedParsing()
        {
            var failed = 0;

            failed += Check("release JSON is read down to the asset list", () =>
            {
                const string json =
                    "{\"tag_name\":\"v1.0.1\",\"name\":\"1.0.1\",\"body\":\"first public build\","
                    + "\"html_url\":\"https://github.com/2mf-hermes/OutlookAiHelper/releases/tag/v1.0.1\","
                    + "\"draft\":false,\"prerelease\":false,\"published_at\":\"2026-10-02T00:00:00Z\","
                    + "\"assets\":[{\"name\":\"OutlookAiHelper.exe\","
                    + "\"browser_download_url\":\"https://github.com/2mf-hermes/OutlookAiHelper/releases/download/v1.0.1/OutlookAiHelper.exe\","
                    + "\"size\":5242880}]}";

                var release = GitHubReleaseFeed.Parse(json);
                return release != null
                    && release.TagName == "v1.0.1"
                    && release.Notes == "first public build"
                    && !release.Draft
                    && release.Assets.Count == 1
                    && release.Assets[0].Name == "OutlookAiHelper.exe"
                    && release.Assets[0].Size == 5242880
                    && release.Assets[0].DownloadUrl.StartsWith("https://github.com/", StringComparison.Ordinal);
            });

            failed += Check("a payload without a tag fails closed", () =>
                GitHubReleaseFeed.Parse("{}") == null
                && GitHubReleaseFeed.Parse(string.Empty) == null
                && GitHubReleaseFeed.Parse("not json at all") == null
                && GitHubReleaseFeed.Parse(null) == null);

            return failed;
        }

        private static int OfferDecision()
        {
            var failed = 0;

            failed += Check("a newer release with a trusted asset may be installed", () =>
            {
                var offer = CheckForUpdateUseCase.Evaluate(
                    Release("v1.0.1", TrustedAssetUrl, 5242880), AppVersion.ParseOrNull("1.0.0"));
                return offer.Status == UpdateStatus.UpdateAvailable
                    && offer.HasUpdate
                    && offer.CanAutoInstall
                    && !offer.AssetBlocked
                    && offer.CurrentVersion == "1.0.0"
                    && offer.LatestVersion == "1.0.1"
                    && offer.AssetName == "OutlookAiHelper.exe"
                    && offer.AssetBytes == 5242880;
            });

            failed += Check("an asset from an untrusted host is shown but never installed", () =>
            {
                var offer = CheckForUpdateUseCase.Evaluate(
                    Release("v1.0.1", "https://evil.example/OutlookAiHelper.exe", 5242880), AppVersion.ParseOrNull("1.0.0"));
                return offer.HasUpdate
                    && offer.AssetBlocked
                    && !offer.CanAutoInstall
                    && string.IsNullOrEmpty(offer.AssetUrl);
            });

            failed += Check("an oversized asset is refused before any download", () =>
            {
                var offer = CheckForUpdateUseCase.Evaluate(
                    Release("v1.0.1", TrustedAssetUrl, UpdateSecurityPolicy.MaxAssetBytes + 1),
                    AppVersion.ParseOrNull("1.0.0"));
                return offer.HasUpdate && offer.AssetBlocked && !offer.CanAutoInstall;
            });

            failed += Check("running the newest build produces no offer", () =>
                CheckForUpdateUseCase.Evaluate(Release("v1.0.0", TrustedAssetUrl, 5242880), AppVersion.ParseOrNull("1.0.0")).Status
                    == UpdateStatus.UpToDate);

            failed += Check("a newer release with no usable asset still points at its page", () =>
            {
                var release = new ReleaseInfo
                {
                    TagName = "v1.0.1",
                    PageUrl = "https://github.com/2mf-hermes/OutlookAiHelper/releases/tag/v1.0.1"
                };
                var offer = CheckForUpdateUseCase.Evaluate(release, AppVersion.ParseOrNull("1.0.0"));
                return offer.HasUpdate
                    && !offer.CanAutoInstall
                    && !offer.AssetBlocked
                    && offer.ReleasePageUrl == "https://github.com/2mf-hermes/OutlookAiHelper/releases/tag/v1.0.1";
            });

            failed += Check("drafts stay invisible and junk tags fail closed", () =>
            {
                var draft = Release("v9.9.9", TrustedAssetUrl, 5242880);
                draft.Draft = true;
                var draftOffer = CheckForUpdateUseCase.Evaluate(draft, AppVersion.ParseOrNull("1.0.0"));

                var junk = CheckForUpdateUseCase.Evaluate(Release("nightly", TrustedAssetUrl, 5242880), AppVersion.ParseOrNull("1.0.0"));

                return draftOffer.Status == UpdateStatus.NotPublished
                    && junk.Status == UpdateStatus.Failed
                    && junk.ErrorKey == "update.error.badRelease"
                    && !junk.CanAutoInstall;
            });

            return failed;
        }

        private static int FeedFailureHandling()
        {
            var failed = 0;

            failed += Check("an unreachable feed never turns into an install", () =>
            {
                var offer = new CheckForUpdateUseCase(new FailingFeed("update.error.timeout"))
                    .Execute(AppVersion.ParseOrNull("1.0.0"));
                return offer.Status == UpdateStatus.Failed
                    && offer.ErrorKey == "update.error.timeout"
                    && !offer.CanAutoInstall;
            });

            failed += Check("a repository without releases is not an error", () =>
            {
                var offer = new CheckForUpdateUseCase(new EmptyFeed()).Execute(AppVersion.ParseOrNull("1.0.0"));
                return offer.Status == UpdateStatus.NotPublished && !offer.CanAutoInstall;
            });

            failed += Check("a missing feed is reported instead of crashing", () =>
                new CheckForUpdateUseCase(null).Execute(AppVersion.ParseOrNull("1.0.0")).ErrorKey == "update.error.config");

            failed += Check("the running version comes from the build stamp", () =>
                !string.IsNullOrEmpty(AppInfo.Version)
                && AppInfo.Version != AppInfo.FallbackVersion
                && AppInfo.CurrentVersion() != null
                && AppInfo.ReleasePageUrl.StartsWith("https://github.com/", StringComparison.Ordinal));

            return failed;
        }

        private const string TrustedAssetUrl =
            "https://github.com/2mf-hermes/OutlookAiHelper/releases/download/v1.0.1/OutlookAiHelper.exe";

        private static ReleaseInfo Release(string tag, string assetUrl, long size)
        {
            var release = new ReleaseInfo
            {
                TagName = tag,
                Name = tag,
                Notes = "notes",
                PageUrl = "https://github.com/2mf-hermes/OutlookAiHelper/releases"
            };
            release.Assets.Add(new ReleaseAsset { Name = "OutlookAiHelper.exe", DownloadUrl = assetUrl, Size = size });
            return release;
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

        private sealed class FailingFeed : IReleaseFeed
        {
            private readonly string _errorKey;

            public FailingFeed(string errorKey)
            {
                _errorKey = errorKey;
            }

            public ReleaseCheckResult FetchLatest(string owner, string repository)
            {
                return ReleaseCheckResult.Failure(_errorKey);
            }
        }

        private sealed class EmptyFeed : IReleaseFeed
        {
            public ReleaseCheckResult FetchLatest(string owner, string repository)
            {
                return ReleaseCheckResult.NoRelease();
            }
        }
    }
}
