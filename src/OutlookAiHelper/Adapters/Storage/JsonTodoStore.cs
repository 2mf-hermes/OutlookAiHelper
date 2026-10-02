using System;
using System.Collections.Generic;
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
                SchemaVersion = 1,
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
                DueOn = string.IsNullOrEmpty(record.DueOn) ? (DateTime?)null : ParseDate(record.DueOn)
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
                DueOn = item.DueOn.HasValue ? item.DueOn.Value.ToString("o") : null
            };
        }

        private static DateTime ParseDate(string value)
        {
            DateTime parsed;
            if (!string.IsNullOrEmpty(value) && DateTime.TryParse(value, out parsed))
            {
                return parsed;
            }

            return DateTime.UtcNow;
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
