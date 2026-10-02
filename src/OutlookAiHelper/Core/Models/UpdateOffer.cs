using System;

namespace OutlookAiHelper.Core.Models
{
    public enum UpdateStatus
    {
        /// <summary>Running build is the newest published release.</summary>
        UpToDate,

        /// <summary>A newer release exists (with or without a downloadable asset).</summary>
        UpdateAvailable,

        /// <summary>The repository has no published release yet.</summary>
        NotPublished,

        /// <summary>The check itself failed (offline, bad payload, blocked source).</summary>
        Failed
    }

    /// <summary>
    /// What the UI is allowed to do after a check. Carrying the decision in one object
    /// keeps "should we download?" out of the window code and inside testable logic.
    /// </summary>
    public sealed class UpdateOffer
    {
        public UpdateStatus Status { get; set; }

        public string CurrentVersion { get; set; }

        public string LatestVersion { get; set; }

        public string ReleaseName { get; set; }

        public string Notes { get; set; }

        public string ReleasePageUrl { get; set; }

        public string AssetName { get; set; }

        public string AssetUrl { get; set; }

        public long AssetBytes { get; set; }

        /// <summary>True when a release asset existed but failed the security policy.</summary>
        public bool AssetBlocked { get; set; }

        public bool Prerelease { get; set; }

        public string ErrorKey { get; set; }

        public bool HasUpdate
        {
            get { return Status == UpdateStatus.UpdateAvailable; }
        }

        /// <summary>Only a policy-approved .exe from a trusted host may be auto-installed.</summary>
        public bool CanAutoInstall
        {
            get
            {
                return Status == UpdateStatus.UpdateAvailable
                    && !AssetBlocked
                    && !string.IsNullOrEmpty(AssetUrl)
                    && !string.IsNullOrEmpty(AssetName);
            }
        }

        public static UpdateOffer UpToDate(string currentVersion)
        {
            return new UpdateOffer
            {
                Status = UpdateStatus.UpToDate,
                CurrentVersion = currentVersion,
                LatestVersion = currentVersion
            };
        }

        public static UpdateOffer NotPublished(string currentVersion)
        {
            return new UpdateOffer
            {
                Status = UpdateStatus.NotPublished,
                CurrentVersion = currentVersion
            };
        }

        public static UpdateOffer Failed(string currentVersion, string errorKey)
        {
            return new UpdateOffer
            {
                Status = UpdateStatus.Failed,
                CurrentVersion = currentVersion,
                ErrorKey = errorKey
            };
        }
    }
}
