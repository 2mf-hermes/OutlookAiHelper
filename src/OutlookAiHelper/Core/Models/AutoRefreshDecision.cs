using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Decides whether the automatic refresh should pay for a full scan.
    ///
    /// The probe is an optimisation, not a precondition. When it cannot answer at all
    /// (Outlook was closed mid-tick, or it failed to attach on the probing thread) the list
    /// still has to refresh eventually, so the ceiling alone decides. A working probe can
    /// only make a scan happen sooner; it must never be able to switch the feature off.
    /// </summary>
    public static class AutoRefreshDecision
    {
        public static bool ShouldScan(
            bool probeUsable,
            bool probeChanged,
            DateTime? lastScanUtc,
            DateTime nowUtc,
            TimeSpan ceiling)
        {
            if (probeUsable && probeChanged)
            {
                return true;
            }

            return !lastScanUtc.HasValue || nowUtc - lastScanUtc.Value >= ceiling;
        }
    }
}
