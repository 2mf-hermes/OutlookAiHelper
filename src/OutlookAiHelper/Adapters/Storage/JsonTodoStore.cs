using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Storage
{
    [DataContract]
    public sealed class TodoDocument
    {
        [DataMember(Order = 1)]
        public int SchemaVersion { get; set; }

        [DataMember(Order = 2)]
        public List<TodoRecord> Items { get; set; }
    }

    [DataContract]
    public sealed class TodoRecord
    {
        [DataMember(Order = 1)]
        public string Id { get; set; }

        [DataMember(Order = 2)]
        public string Title { get; set; }

        [DataMember(Order = 3)]
        public string Note { get; set; }

        [DataMember(Order = 4)]
        public string SourceEntryId { get; set; }

        [DataMember(Order = 5)]
        public string QuadrantHint { get; set; }

        [DataMember(Order = 6)]
        public string Status { get; set; }

        [DataMember(Order = 7)]
        public string CreatedOn { get; set; }

        [DataMember(Order = 8)]
        public string UpdatedOn { get; set; }

        [DataMember(Order = 9)]
        public string DueOn { get; set; }

        [DataMember(Order = 10)]
        public string SourceFrom { get; set; }

        // Stored as a round-trip UTC timestamp. Absent in files written before this
        // field existed, which deserializes to null and falls back to CreatedOn.
        [DataMember(Order = 11)]
        public string SourceReceivedOn { get; set; }

        // Added in schema 2. Every one of these is optional: a file written by an
        // earlier build simply has no member and deserializes to null/0.
        [DataMember(Order = 12)]
        public string ReminderOn { get; set; }

        [DataMember(Order = 13)]
        public string Repeat { get; set; }

        [DataMember(Order = 14)]
        public int RepeatAnchorDay { get; set; }

        [DataMember(Order = 15)]
        public string Category { get; set; }

        [DataMember(Order = 16)]
        public List<TodoAttachmentRecord> Attachments { get; set; }
    }

    /// <summary>
    /// The metadata of one attached file. Its bytes are not in this document; the
    /// record points at a file the attachment store copied into its own folder.
    /// </summary>
    [DataContract]
    public sealed class TodoAttachmentRecord
    {
        [DataMember(Order = 1)]
        public string FileName { get; set; }

        [DataMember(Order = 2)]
        public string StoredName { get; set; }

        [DataMember(Order = 3)]
        public long SizeBytes { get; set; }

        [DataMember(Order = 4)]
        public string AddedOn { get; set; }
    }

    public sealed class JsonTodoStore : ITodoStore
    {
        private readonly string _path;

        public JsonTodoStore()
            : this(Path.Combine(JsonPaths.RootDirectory, "todos.json"))
        {
        }

        public JsonTodoStore(string path)
        {
            _path = path;
        }

        public IList<TodoItem> Load()
        {
            var list = new List<TodoItem>();
            if (!File.Exists(_path))
            {
                return list;
            }

            try
            {
                using (var stream = File.OpenRead(_path))
                {
                    var doc = (TodoDocument)new DataContractJsonSerializer(typeof(TodoDocument)).ReadObject(stream);
                    if (doc == null || doc.Items == null)
                    {
                        return list;
                    }

                    foreach (var record in doc.Items)
                    {
                        if (record == null || string.IsNullOrEmpty(record.Id))
                        {
                            continue;
                        }

                        list.Add(ToItem(record));
                    }
                }
            }
            catch (Exception)
            {
                TryQuarantine();
            }

            return list;
        }

        public void Save(IList<TodoItem> items)
        {
            JsonPaths.EnsureRoot();
            var doc = new TodoDocument
            {
                SchemaVersion = 2,
                Items = new List<TodoRecord>()
            };

            if (items != null)
            {
                foreach (var item in items)
                {
                    if (item == null)
                    {
                        continue;
                    }

                    doc.Items.Add(ToRecord(item));
                }
            }

            using (var stream = File.Create(_path))
            {
                new DataContractJsonSerializer(typeof(TodoDocument)).WriteObject(stream, doc);
            }
        }

        private static TodoItem ToItem(TodoRecord record)
        {
            var status = TodoStatus.Open;
            if (string.Equals(record.Status, "Done", StringComparison.OrdinalIgnoreCase))
            {
                status = TodoStatus.Done;
            }

            return new TodoItem
            {
                Id = record.Id,
                Title = record.Title ?? string.Empty,
                Note = record.Note ?? string.Empty,
                SourceEntryId = record.SourceEntryId,
                QuadrantHint = record.QuadrantHint,
                Status = status,
                CreatedOn = ParseDate(record.CreatedOn),
                UpdatedOn = ParseDate(record.UpdatedOn),
                DueOn = ParseDateOrNull(record.DueOn),
                ReminderOn = ParseDateOrNull(record.ReminderOn),
                Repeat = TodoRecurrence.Parse(record.Repeat),
                RepeatAnchorDay = record.RepeatAnchorDay,
                Category = record.Category ?? string.Empty,
                Attachments = ToAttachments(record.Attachments),
                SourceFrom = record.SourceFrom,
                SourceReceivedOn = string.IsNullOrEmpty(record.SourceReceivedOn)
                    ? (DateTime?)null
                    : ParseDate(record.SourceReceivedOn)
            };
        }

        private static TodoRecord ToRecord(TodoItem item)
        {
            return new TodoRecord
            {
                Id = item.Id,
                Title = item.Title,
                Note = item.Note,
                SourceEntryId = item.SourceEntryId,
                QuadrantHint = item.QuadrantHint,
                Status = item.Status == TodoStatus.Done ? "Done" : "Open",
                CreatedOn = item.CreatedOn.ToString("o"),
                UpdatedOn = item.UpdatedOn.ToString("o"),
                DueOn = item.DueOn.HasValue ? item.DueOn.Value.ToString("o") : null,
                ReminderOn = item.ReminderOn.HasValue ? item.ReminderOn.Value.ToString("o") : null,
                Repeat = TodoRecurrence.Key(item.Repeat),
                RepeatAnchorDay = item.RepeatAnchorDay,
                Category = item.Category,
                Attachments = ToAttachmentRecords(item.Attachments),
                SourceFrom = item.SourceFrom,
                SourceReceivedOn = item.SourceReceivedOn.HasValue
                    ? item.SourceReceivedOn.Value.ToString("o")
                    : null
            };
        }

        private static DateTime ParseDate(string value)
        {
            DateTime parsed;
            // RoundtripKind keeps the stored UTC instant intact. Plain TryParse would
            // shift a "…Z" value into local time, which silently makes the overdue and
            // "waiting N days" comparisons wrong by the timezone offset.
            if (!string.IsNullOrEmpty(value)
                && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed))
            {
                return parsed;
            }

            return DateTime.UtcNow;
        }

        /// <summary>
        /// A date that is allowed to be absent, and that is absent when the stored text
        /// is unreadable too. <see cref="ParseDate"/> falls back to "now", which is
        /// right for CreatedOn but wrong for a deadline: a corrupt value must not turn
        /// a follow-up into an overdue one.
        /// </summary>
        private static DateTime? ParseDateOrNull(string value)
        {
            DateTime parsed;
            if (!string.IsNullOrEmpty(value)
                && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed))
            {
                return parsed;
            }

            return null;
        }

        private static List<TodoAttachment> ToAttachments(List<TodoAttachmentRecord> records)
        {
            var list = new List<TodoAttachment>();
            if (records == null)
            {
                return list;
            }

            foreach (var record in records)
            {
                if (record == null || string.IsNullOrEmpty(record.StoredName))
                {
                    continue;
                }

                list.Add(new TodoAttachment
                {
                    FileName = record.FileName ?? string.Empty,
                    StoredName = record.StoredName,
                    SizeBytes = record.SizeBytes,
                    AddedOn = ParseDateOrNull(record.AddedOn) ?? DateTime.UtcNow
                });
            }

            return list;
        }

        private static List<TodoAttachmentRecord> ToAttachmentRecords(List<TodoAttachment> attachments)
        {
            var list = new List<TodoAttachmentRecord>();
            if (attachments == null)
            {
                return list;
            }

            foreach (var attachment in attachments)
            {
                if (attachment == null || string.IsNullOrEmpty(attachment.StoredName))
                {
                    continue;
                }

                list.Add(new TodoAttachmentRecord
                {
                    FileName = attachment.FileName,
                    StoredName = attachment.StoredName,
                    SizeBytes = attachment.SizeBytes,
                    AddedOn = attachment.AddedOn.ToString("o")
                });
            }

            return list;
        }

        private void TryQuarantine()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    return;
                }

                File.Move(_path, _path + ".bad-" + DateTime.Now.ToString("yyyyMMddHHmmss"));
            }
            catch (Exception)
            {
            }
        }
    }
}
