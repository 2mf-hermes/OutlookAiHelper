using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Storage
{
    public static class JsonPaths
    {
        public static string RootDirectory
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAiHelper"); }
        }

        public static string OverridesPath
        {
            get { return Path.Combine(RootDirectory, "overrides.json"); }
        }

        public static string SettingsPath
        {
            get { return Path.Combine(RootDirectory, "settings.json"); }
        }

        public static void EnsureRoot()
        {
            Directory.CreateDirectory(RootDirectory);
        }
    }

    [DataContract]
    public sealed class OverrideDocument
    {
        [DataMember(Order = 1)]
        public int SchemaVersion { get; set; }

        [DataMember(Order = 2)]
        public List<OverrideRecord> Items { get; set; }
    }

    [DataContract]
    public sealed class OverrideRecord
    {
        [DataMember(Order = 1)]
        public string EntryId { get; set; }

        [DataMember(Order = 2)]
        public string Quadrant { get; set; }

        [DataMember(Order = 3)]
        public string UpdatedOn { get; set; }

        [DataMember(Order = 4)]
        public string Source { get; set; }
    }

    public sealed class JsonOverrideStore : IOverrideStore
    {
        private readonly string _path;

        public JsonOverrideStore()
            : this(JsonPaths.OverridesPath)
        {
        }

        public JsonOverrideStore(string path)
        {
            _path = path;
        }

        public IDictionary<string, OverrideEntry> Load()
        {
            var map = new Dictionary<string, OverrideEntry>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(_path))
            {
                return map;
            }

            try
            {
                using (var stream = File.OpenRead(_path))
                {
                    var doc = (OverrideDocument)ReadObject(typeof(OverrideDocument), stream);
                    if (doc == null || doc.Items == null)
                    {
                        return map;
                    }

                    foreach (var item in doc.Items)
                    {
                        if (item == null || string.IsNullOrEmpty(item.EntryId))
                        {
                            continue;
                        }

                        Quadrant quadrant;
                        if (!Enum.TryParse(item.Quadrant, true, out quadrant))
                        {
                            continue;
                        }

                        DateTime updatedOn;
                        if (!DateTime.TryParse(item.UpdatedOn, out updatedOn))
                        {
                            updatedOn = DateTime.UtcNow;
                        }

                        map[item.EntryId] = new OverrideEntry
                        {
                            EntryId = item.EntryId,
                            Quadrant = quadrant,
                            UpdatedOn = updatedOn,
                            Source = item.Source
                        };
                    }
                }
            }
            catch (Exception)
            {
                TryQuarantine();
            }

            return map;
        }

        public void Save(IDictionary<string, OverrideEntry> overrides)
        {
            JsonPaths.EnsureRoot();
            var doc = new OverrideDocument
            {
                SchemaVersion = 1,
                Items = new List<OverrideRecord>()
            };

            if (overrides != null)
            {
                foreach (var pair in overrides)
                {
                    if (pair.Value == null)
                    {
                        continue;
                    }

                    doc.Items.Add(new OverrideRecord
                    {
                        EntryId = pair.Key,
                        Quadrant = pair.Value.Quadrant.ToString(),
                        UpdatedOn = pair.Value.UpdatedOn.ToString("o"),
                        Source = pair.Value.Source
                    });
                }
            }

            using (var stream = File.Create(_path))
            {
                WriteObject(typeof(OverrideDocument), stream, doc);
            }
        }

        private void TryQuarantine()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    return;
                }

                var bad = _path + ".bad-" + DateTime.Now.ToString("yyyyMMddHHmmss");
                File.Move(_path, bad);
            }
            catch (Exception)
            {
            }
        }

        private static object ReadObject(Type type, Stream stream)
        {
            var serializer = new DataContractJsonSerializer(type);
            return serializer.ReadObject(stream);
        }

        private static void WriteObject(Type type, Stream stream, object graph)
        {
            var serializer = new DataContractJsonSerializer(type);
            serializer.WriteObject(stream, graph);
        }
    }
}
