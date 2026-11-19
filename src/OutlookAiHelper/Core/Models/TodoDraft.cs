using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// What the add/edit sheet collects. The window fills it, the main window applies
    /// it (copying files, saving the item), which keeps the sheet free of storage
    /// concerns and lets the tests drive the same data without a UI.
    /// </summary>
    public sealed class TodoDraft
    {
        public TodoDraft()
        {
            NewFiles = new List<string>();
            RemovedStoredNames = new List<string>();
        }

        public string Title { get; set; }
        public string Note { get; set; }
        public string Category { get; set; }
        public DateTime? DueOn { get; set; }
        public DateTime? ReminderOn { get; set; }
        public TodoRepeat Repeat { get; set; }
        public int RepeatAnchorDay { get; set; }

        /// <summary>Null when the follow-up is not (or no longer) linked to a mail.</summary>
        public string SourceEntryId { get; set; }
        public string SourceFrom { get; set; }
        public DateTime? SourceReceivedOn { get; set; }

        /// <summary>Full paths the user picked in this session; copied in on save.</summary>
        public List<string> NewFiles { get; private set; }

        /// <summary>Stored names the user removed in the sheet; deleted on save.</summary>
        public List<string> RemovedStoredNames { get; private set; }
    }
}
