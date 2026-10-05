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
    public sealed class SettingsDocument
    {
        [DataMember(Order = 1)]
        public int SchemaVersion { get; set; }

        [DataMember(Order = 2)]
        public string Language { get; set; }

        [DataMember(Order = 3)]
        public int ScanDays { get; set; }

        [DataMember(Order = 4)]
        public string FolderPath { get; set; }

        [DataMember(Order = 5)]
        public List<string> VipAddresses { get; set; }

        [DataMember(Order = 6)]
        public List<string> UrgentKeywords { get; set; }

        [DataMember(Order = 7)]
        public List<string> ImportantKeywords { get; set; }

        [DataMember(Order = 8)]
        public bool AiEnabled { get; set; }

        [DataMember(Order = 9)]
        public string AiBaseUrl { get; set; }

        [DataMember(Order = 10)]
        public string AiModel { get; set; }

        [DataMember(Order = 11)]
        public string AiApiKey { get; set; }

        [DataMember(Order = 12)]
        public List<ProviderRecord> AiProviders { get; set; }

        [DataMember(Order = 13)]
        public string SelectedAiProviderId { get; set; }

        /// <summary>Opt-in update check. Absent in older files, which reads as false.</summary>
        [DataMember(Order = 14)]
        public bool CheckForUpdatesOnStartup { get; set; }

        /// <summary>
        /// Inbox probe interval in minutes. Nullable so that "field absent" (an older
        /// settings file) can be told apart from the user's explicit 0 = off; an int
        /// would collapse both to zero and silently disable the feature for everyone
        /// who upgrades.
        /// </summary>
        [DataMember(Order = 15)]
        public int? AutoRefreshMinutes { get; set; }
    }

    [DataContract]
    public sealed class ProviderRecord
    {
        [DataMember(Order = 1)]
        public string Id { get; set; }

        [DataMember(Order = 2)]
        public string Name { get; set; }

        [DataMember(Order = 3)]
        public string BaseUrl { get; set; }

        [DataMember(Order = 4)]
        public string ApiKey { get; set; }

        [DataMember(Order = 5)]
        public string Model { get; set; }

        [DataMember(Order = 6)]
        public List<string> CachedModels { get; set; }
    }

    public sealed class JsonSettingsStore : ISettingsStore
    {
        private readonly string _path;

        public JsonSettingsStore()
            : this(JsonPaths.SettingsPath)
        {
        }

        public JsonSettingsStore(string path)
        {
            _path = path;
        }

        public AppSettings Load()
        {
            if (!File.Exists(_path))
            {
                return AppSettings.CreateDefault();
            }

            try
            {
                using (var stream = File.OpenRead(_path))
                {
                    var doc = (SettingsDocument)new DataContractJsonSerializer(typeof(SettingsDocument)).ReadObject(stream);
                    if (doc == null)
                    {
                        return AppSettings.CreateDefault();
                    }

                    return new AppSettings
                    {
                        SchemaVersion = doc.SchemaVersion <= 0 ? 1 : doc.SchemaVersion,
                        Language = string.IsNullOrEmpty(doc.Language) ? "zh-TW" : doc.Language,
                        ScanDays = doc.ScanDays <= 0 ? 30 : doc.ScanDays,
                        FolderPath = string.IsNullOrEmpty(doc.FolderPath) ? "Inbox" : doc.FolderPath,
                        VipAddresses = doc.VipAddresses ?? new List<string>(),
                        UrgentKeywords = doc.UrgentKeywords != null && doc.UrgentKeywords.Count > 0
                            ? doc.UrgentKeywords
                            : Core.Classification.RuleOptions.DefaultUrgentKeywords(),
                        ImportantKeywords = doc.ImportantKeywords != null && doc.ImportantKeywords.Count > 0
                            ? doc.ImportantKeywords
                            : Core.Classification.RuleOptions.DefaultImportantKeywords(),
                        AiEnabled = doc.AiEnabled,
                        AiBaseUrl = doc.AiBaseUrl ?? string.Empty,
                        AiModel = doc.AiModel ?? string.Empty,
                        AiApiKey = doc.AiApiKey ?? string.Empty,
                        AiProviders = FromProviderRecords(doc.AiProviders),
                        SelectedAiProviderId = doc.SelectedAiProviderId ?? string.Empty,
                        CheckForUpdatesOnStartup = doc.CheckForUpdatesOnStartup,
                        AutoRefreshMinutes = NormalizeAutoRefreshMinutes(doc.AutoRefreshMinutes)
                    };
                }
            }
            catch (Exception)
            {
                TryQuarantine();
                return AppSettings.CreateDefault();
            }
        }

        public void Save(AppSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            JsonPaths.EnsureRoot();
            var doc = new SettingsDocument
            {
                SchemaVersion = 1,
                Language = settings.Language,
                ScanDays = settings.ScanDays,
                FolderPath = settings.FolderPath,
                VipAddresses = settings.VipAddresses ?? new List<string>(),
                UrgentKeywords = settings.UrgentKeywords ?? new List<string>(),
                ImportantKeywords = settings.ImportantKeywords ?? new List<string>(),
                AiEnabled = settings.AiEnabled,
                AiBaseUrl = settings.AiBaseUrl ?? string.Empty,
                AiModel = settings.AiModel ?? string.Empty,
                AiApiKey = settings.AiApiKey ?? string.Empty,
                AiProviders = ToProviderRecords(settings.AiProviders),
                SelectedAiProviderId = settings.SelectedAiProviderId ?? string.Empty,
                CheckForUpdatesOnStartup = settings.CheckForUpdatesOnStartup,
                AutoRefreshMinutes = NormalizeAutoRefreshMinutes(settings.AutoRefreshMinutes)
            };

            using (var stream = File.Create(_path))
            {
                new DataContractJsonSerializer(typeof(SettingsDocument)).WriteObject(stream, doc);
            }
        }

        /// <summary>
        /// Applies the probe-interval rules in one place: a missing field keeps the app
        /// default (2 minutes), and anything outside 0..30 — including a hand-edited file
        /// — falls back to the default rather than to "off", so a corrupt value cannot
        /// quietly switch the feature off.
        /// </summary>
        private static int NormalizeAutoRefreshMinutes(int? value)
        {
            if (!value.HasValue)
            {
                return AppSettings.DefaultAutoRefreshMinutes;
            }

            var minutes = value.Value;
            if (minutes < 0 || minutes > AppSettings.MaxAutoRefreshMinutes)
            {
                return AppSettings.DefaultAutoRefreshMinutes;
            }

            return minutes;
        }

        private static List<AiProviderProfile> FromProviderRecords(List<ProviderRecord> records)
        {
            var list = new List<AiProviderProfile>();
            if (records == null)
            {
                return list;
            }

            foreach (var r in records)
            {
                if (r == null || string.IsNullOrEmpty(r.Id))
                {
                    continue;
                }

                list.Add(new AiProviderProfile
                {
                    Id = r.Id,
                    Name = r.Name ?? "Provider",
                    BaseUrl = r.BaseUrl ?? string.Empty,
                    ApiKey = r.ApiKey ?? string.Empty,
                    Model = r.Model ?? string.Empty,
                    CachedModels = r.CachedModels ?? new List<string>()
                });
            }

            return list;
        }

        private static List<ProviderRecord> ToProviderRecords(List<AiProviderProfile> profiles)
        {
            var list = new List<ProviderRecord>();
            if (profiles == null)
            {
                return list;
            }

            foreach (var item in profiles)
            {
                if (item == null)
                {
                    continue;
                }

                list.Add(new ProviderRecord
                {
                    Id = item.Id,
                    Name = item.Name,
                    BaseUrl = item.BaseUrl,
                    ApiKey = item.ApiKey,
                    Model = item.Model,
                    CachedModels = item.CachedModels ?? new List<string>()
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
