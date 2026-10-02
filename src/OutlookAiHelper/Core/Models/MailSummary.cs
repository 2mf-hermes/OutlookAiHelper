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
    }
}
