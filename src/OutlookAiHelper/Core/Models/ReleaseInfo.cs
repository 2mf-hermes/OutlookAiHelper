using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>One downloadable file attached to a GitHub release.</summary>
    public sealed class ReleaseAsset
    {
        public string Name { get; set; }

        public string DownloadUrl { get; set; }

        public long Size { get; set; }
    }

    /// <summary>Trimmed view of a GitHub release — only what the update check needs.</summary>
    public sealed class ReleaseInfo
    {
        public ReleaseInfo()
        {
            Assets = new List<ReleaseAsset>();
        }

        public string TagName { get; set; }

        public string Name { get; set; }

        public string Notes { get; set; }

        public string PageUrl { get; set; }

        public bool Draft { get; set; }

        public bool Prerelease { get; set; }

        public string PublishedAt { get; set; }

        public List<ReleaseAsset> Assets { get; set; }
    }

    /// <summary>Outcome of asking the release feed for the newest release.</summary>
    public sealed class ReleaseCheckResult
    {
        public bool Ok { get; set; }

        public bool NothingPublished { get; set; }

        public string ErrorKey { get; set; }

        public ReleaseInfo Release { get; set; }

        public static ReleaseCheckResult Success(ReleaseInfo release)
        {
            return new ReleaseCheckResult { Ok = true, Release = release };
        }

        public static ReleaseCheckResult NoRelease()
        {
            return new ReleaseCheckResult { Ok = true, NothingPublished = true };
        }

        public static ReleaseCheckResult Failure(string errorKey)
        {
            return new ReleaseCheckResult { Ok = false, ErrorKey = errorKey };
        }
    }
}
