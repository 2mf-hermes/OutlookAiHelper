using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Result of the lightweight "did anything arrive?" probe the auto-refresh timer
    /// runs between full scans.
    ///
    /// Why a probe at all: a full scan walks every folder of every store and re-reads
    /// up to two thousand items. A check that reads only the head of each Inbox costs
    /// a fraction of that, so the app can tell the difference between "nothing changed"
    /// and "new mail arrived". Only the second case is worth a rescan.
    ///
    /// Deliberately read-only: nothing here sorts or mutates a folder view, and no
    /// message is touched — the probe never marks anything read.
    /// </summary>
    public sealed class InboxProbe
    {
        /// <summary>
        /// Changed when the newest item at the head of any Inbox changes (new mail,
        /// deletion, or a move) or when the unread total moves (covers mail a rule
        /// files straight out of the Inbox, and read/unread flips made in Outlook).
        /// </summary>
        public string Signature { get; set; }

        /// <summary>Unread items across the probed Inboxes; shown in the status line.</summary>
        public int UnreadCount { get; set; }

        /// <summary>
        /// Inboxes actually read. Zero means no store answered (Outlook closed mid-poll),
        /// in which case the caller must skip this tick rather than conclude "no change".
        /// </summary>
        public int InboxCount { get; set; }

        public bool IsUsable
        {
            get { return InboxCount > 0 && !string.IsNullOrEmpty(Signature); }
        }

        /// <summary>Builds the per-Inbox signature fragment from one Inbox head.</summary>
        public static string MakeFragment(string newestEntryId, DateTime receivedOn)
        {
            return (string.IsNullOrEmpty(newestEntryId) ? "-" : newestEntryId)
                + "@" + receivedOn.ToUniversalTime().Ticks;
        }
    }
}
