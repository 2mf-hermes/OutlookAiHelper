using System;

namespace OutlookAiHelper.Core.Models
{
    public enum TodoStatus
    {
        Open,
        Done
    }

    public sealed class TodoItem
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Note { get; set; }
        public string SourceEntryId { get; set; }
        public string QuadrantHint { get; set; }
        public TodoStatus Status { get; set; }
        public DateTime CreatedOn { get; set; }
        public DateTime UpdatedOn { get; set; }
        public DateTime? DueOn { get; set; }

        /// <summary>
        /// Snapshot of the source mail's sender. Kept as text so the follow-up stays
        /// readable after the mail is deleted, moved or the store is disconnected.
        /// </summary>
        public string SourceFrom { get; set; }

        /// <summary>
        /// When the source mail arrived, as a UTC instant, so "waiting N days" survives
        /// a timezone change and older stored items (null) fall back to CreatedOn.
        /// </summary>
        public DateTime? SourceReceivedOn { get; set; }

        /// <summary>
        /// How long this item has been waiting for an answer: from the source mail's
        /// received time when known, otherwise since the item was created. Never
        /// negative, so a clock skew on the mail server cannot print "-1 天".
        /// </summary>
        public int WaitingDays(DateTime nowUtc)
        {
            var since = SourceReceivedOn ?? CreatedOn;
            var days = (int)Math.Floor((nowUtc - since).TotalDays);
            return days < 0 ? 0 : days;
        }

        /// <summary>Open items only: a finished follow-up is never "overdue".</summary>
        public bool IsOverdue(DateTime nowUtc)
        {
            return Status == TodoStatus.Open && DueOn.HasValue && DueOn.Value < nowUtc;
        }

        public static TodoItem Create(
            string title,
            string sourceEntryId = null,
            string quadrantHint = null,
            string sourceFrom = null,
            DateTime? sourceReceivedOn = null)
        {
            var now = DateTime.UtcNow;
            return new TodoItem
            {
                Id = Guid.NewGuid().ToString("N"),
                Title = title ?? string.Empty,
                Note = string.Empty,
                SourceEntryId = sourceEntryId,
                QuadrantHint = quadrantHint,
                SourceFrom = sourceFrom,
                SourceReceivedOn = sourceReceivedOn,
                Status = TodoStatus.Open,
                CreatedOn = now,
                UpdatedOn = now
            };
        }
    }
}
