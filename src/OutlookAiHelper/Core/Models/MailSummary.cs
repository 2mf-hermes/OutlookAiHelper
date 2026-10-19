using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    public sealed class MailSummary
    {
        public string EntryId { get; set; }
        public string FolderPath { get; set; }
        public string Subject { get; set; }
        public string FromName { get; set; }
        public string FromAddress { get; set; }
        public DateTime ReceivedOn { get; set; }
        public bool IsRead { get; set; }
        public bool HasFlag { get; set; }
        public bool IsMeetingRequest { get; set; }
        public int Importance { get; set; }

        /// <summary>
        /// The opening lines of the mail's own text, flattened to one line and cut short, so
        /// two mails with the same subject can still be told apart in a list. Display only:
        /// never stored, logged or sent. Null when the source cannot supply it - the cloud
        /// reader only asks for Mail.ReadBasic, which does not include a body.
        /// </summary>
        public string Preview { get; set; }
    }
}
