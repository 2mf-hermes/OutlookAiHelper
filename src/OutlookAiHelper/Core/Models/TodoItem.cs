using System;
using System.Collections.Generic;

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
        /// When the follow-up should be brought back to the user's attention. Its own
        /// field rather than a copy of <see cref="DueOn"/>, so a reminder can be switched
        /// off while the deadline stays, which is how the mainstream task apps behave.
        /// </summary>
        public DateTime? ReminderOn { get; set; }

        /// <summary>How the item comes back after it is ticked off; None is a one-off.</summary>
        public TodoRepeat Repeat { get; set; }

        /// <summary>
        /// Day of the month a monthly/yearly series is anchored on (0 = the due date's own
        /// day). Without it, 1 月 31 日 would collapse to 2 月 28 日 and stay there.
        /// </summary>
        public int RepeatAnchorDay { get; set; }

        /// <summary>Free-text category, empty when the user gave none.</summary>
        public string Category { get; set; }

        /// <summary>
        /// Copies of the files the follow-up carries. The bytes live outside the item
        /// document (see TodoAttachmentStore); an older file without this member
        /// deserializes to null, so every reader treats null as empty.
        /// </summary>
        public List<TodoAttachment> Attachments { get; set; }

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

        /// <summary>
        /// True when the reminder is due and the item is still open. A reminder in the
        /// future, an acknowledged or a finished item never counts.
        /// </summary>
        public bool IsReminderDue(DateTime nowUtc)
        {
            return Status == TodoStatus.Open && ReminderOn.HasValue && ReminderOn.Value <= nowUtc;
        }

        /// <summary>Never null, so callers can count without a null check.</summary>
        public int AttachmentCount
        {
            get { return Attachments == null ? 0 : Attachments.Count; }
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
                UpdatedOn = now,
                Repeat = TodoRepeat.None,
                Category = string.Empty,
                Attachments = new List<TodoAttachment>()
            };
        }
    }
}
