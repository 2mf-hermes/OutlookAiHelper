using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public sealed class MailScanScope
    {
        public IList<string> FolderPaths { get; set; }
        public int Days { get; set; }
    }

    /// <summary>
    /// One mailbox, read-only. Implemented once per backing service (the Outlook desktop
    /// client today, Microsoft Graph next), so the scan, the classifier and the whole UI
    /// stay unaware of where a mail came from.
    /// </summary>
    public interface IMailReader
    {
        bool IsAvailable(out string reasonKey);
        IList<string> ListFolders();
        IEnumerable<MailSummary> Enumerate(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled);

        /// <summary>
        /// The cheap "did anything arrive?" check the auto-refresh timer runs between full
        /// scans. Part of the port rather than of the Outlook adapter because the timer
        /// has to be able to ask whichever source is active.
        /// </summary>
        InboxProbe ProbeInbox();

        /// <summary>
        /// Brings one mail up in its native client. False means this source has no such
        /// client (a cloud source has no Outlook window to open), which the caller reports
        /// as "cannot open here" rather than as "mail not found".
        /// </summary>
        bool TryOpenInOutlook(string entryId);
    }
}
