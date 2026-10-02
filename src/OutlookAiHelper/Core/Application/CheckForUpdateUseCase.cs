using System;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Core.Security;

namespace OutlookAiHelper.Core.Application
{
    /// <summary>
    /// Turns "what is the newest release?" into "what may this app do about it?".
    /// Pure decision logic: no network, no file system, no UI.
    /// </summary>
    public sealed class CheckForUpdateUseCase
    {
        private readonly IReleaseFeed _feed;

        public CheckForUpdateUseCase(IReleaseFeed feed)
        {
            _feed = feed;
        }

        public UpdateOffer Execute(AppVersion current)
        {
            if (current == null)
            {
                current = AppVersion.ParseOrNull(AppInfo.FallbackVersion);
            }

            if (_feed == null)
            {
                return UpdateOffer.Failed(current.Normalized, "update.error.config");
            }

            ReleaseCheckResult result;
            try
            {
                result = _feed.FetchLatest(AppInfo.Owner, AppInfo.Repo);
            }
            catch (Exception)
            {
                return UpdateOffer.Failed(current.Normalized, "update.error.network");
            }

            if (result == null)
            {
                return UpdateOffer.Failed(current.Normalized, "update.error.network");
            }

            if (!result.Ok)
            {
                return UpdateOffer.Failed(current.Normalized, result.ErrorKey ?? "update.error.network");
            }

            if (result.NothingPublished || result.Release == null)
            {
                return UpdateOffer.NotPublished(current.Normalized);
            }

            return Evaluate(result.Release, current);
        }

        /// <summary>
        /// Decides what the user may be offered for one release. A newer release whose
        /// asset fails the security policy is still reported — but the offer then only
        /// carries the release page link, never a download.
        /// </summary>
        public static UpdateOffer Evaluate(ReleaseInfo release, AppVersion current)
        {
            if (current == null)
            {
                current = AppVersion.ParseOrNull(AppInfo.FallbackVersion);
            }

            if (release == null || release.Draft)
            {
                return UpdateOffer.NotPublished(current.Normalized);
            }

            var latest = AppVersion.ParseOrNull(release.TagName);
            if (latest == null)
            {
                return UpdateOffer.Failed(current.Normalized, "update.error.badRelease");
            }

            if (!latest.IsNewerThan(current))
            {
                return UpdateOffer.UpToDate(current.Normalized);
            }

            var offer = new UpdateOffer
            {
                Status = UpdateStatus.UpdateAvailable,
                CurrentVersion = current.Normalized,
                LatestVersion = latest.Normalized,
                ReleaseName = string.IsNullOrEmpty(release.Name) ? release.TagName : release.Name,
                Notes = release.Notes,
                ReleasePageUrl = string.IsNullOrEmpty(release.PageUrl) ? AppInfo.ReleasePageUrl : release.PageUrl,
                Prerelease = release.Prerelease
            };

            var asset = UpdateSecurityPolicy.SelectAsset(release);
            if (asset != null)
            {
                if (UpdateSecurityPolicy.IsTrustedAssetUrl(asset.DownloadUrl)
                    && UpdateSecurityPolicy.IsWithinSizeLimit(asset.Size))
                {
                    offer.AssetName = asset.Name.Trim();
                    offer.AssetUrl = asset.DownloadUrl;
                    offer.AssetBytes = asset.Size;
                }
                else
                {
                    // Reported to the UI as "we will not auto-install this".
                    offer.AssetBlocked = true;
                }
            }

            return offer;
        }
    }
}
