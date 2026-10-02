using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    /// <summary>Read-only source of the newest published release.</summary>
    public interface IReleaseFeed
    {
        ReleaseCheckResult FetchLatest(string owner, string repository);
    }
}
