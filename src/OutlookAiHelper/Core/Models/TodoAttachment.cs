using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// One file the user attached to a follow-up. The file itself lives outside the
    /// JSON document - in the per-item folder <c>attachments\{todoId}\</c> under the
    /// app's data folder - so a large PDF never has to be serialised into todos.json.
    /// <see cref="StoredName"/> is the on-disk name; <see cref="FileName"/> is what the
    /// user picked and what the panel shows.
    /// </summary>
    public sealed class TodoAttachment
    {
        public string FileName { get; set; }
        public string StoredName { get; set; }
        public long SizeBytes { get; set; }
        public DateTime AddedOn { get; set; }

        public static TodoAttachment Create(string fileName, string storedName, long sizeBytes)
        {
            return new TodoAttachment
            {
                FileName = fileName ?? string.Empty,
                StoredName = storedName ?? string.Empty,
                SizeBytes = sizeBytes,
                AddedOn = DateTime.UtcNow
            };
        }
    }
}
