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

        public static TodoItem Create(string title, string sourceEntryId = null, string quadrantHint = null)
        {
            var now = DateTime.UtcNow;
            return new TodoItem
            {
                Id = Guid.NewGuid().ToString("N"),
                Title = title ?? string.Empty,
                Note = string.Empty,
                SourceEntryId = sourceEntryId,
                QuadrantHint = quadrantHint,
                Status = TodoStatus.Open,
                CreatedOn = now,
                UpdatedOn = now
            };
        }
    }
}
