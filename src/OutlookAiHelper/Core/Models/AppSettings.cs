using System;
using System.Collections.Generic;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.Core.Models
{
    public sealed class AppSettings
    {
        public AppSettings()
        {
            SchemaVersion = 1;
            Language = "zh-TW";
            ScanDays = 30;
            FolderPath = "Inbox";
            VipAddresses = new List<string>();
            UrgentKeywords = new List<string>();
            ImportantKeywords = new List<string>();
            AiEnabled = false;
            AiBaseUrl = string.Empty;
            AiModel = string.Empty;
            AiApiKey = string.Empty;
            AiProviders = new List<AiProviderProfile>();
            SelectedAiProviderId = string.Empty;
        }

        public int SchemaVersion { get; set; }
        public string Language { get; set; }
        public int ScanDays { get; set; }
        public string FolderPath { get; set; }
        public List<string> VipAddresses { get; set; }
        public List<string> UrgentKeywords { get; set; }
        public List<string> ImportantKeywords { get; set; }
        public bool AiEnabled { get; set; }
        public string AiBaseUrl { get; set; }
        public string AiModel { get; set; }
        public string AiApiKey { get; set; }
        public List<AiProviderProfile> AiProviders { get; set; }
        public string SelectedAiProviderId { get; set; }

        /// <summary>
        /// Opt-in. The update check is the only request this app makes on its own behalf,
        /// so it stays off until the user asks for it; absence in an older settings file
        /// deserializes to false, which is the same answer.
        /// </summary>
        public bool CheckForUpdatesOnStartup { get; set; }

        public AiProviderProfile GetSelectedProvider()
        {
            if (AiProviders == null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(SelectedAiProviderId))
            {
                foreach (var p in AiProviders)
                {
                    if (p != null && p.Id == SelectedAiProviderId)
                    {
                        return p;
                    }
                }
            }

            return AiProviders.Count > 0 ? AiProviders[0] : null;
        }

        public UiLanguage ToUiLanguage()
        {
            if (string.Equals(Language, "zh-CN", StringComparison.OrdinalIgnoreCase))
            {
                return UiLanguage.ZhCn;
            }

            if (string.Equals(Language, "en-US", StringComparison.OrdinalIgnoreCase))
            {
                return UiLanguage.EnUs;
            }

            return UiLanguage.ZhTw;
        }

        public static AppSettings CreateDefault()
        {
            var settings = new AppSettings();
            settings.UrgentKeywords = Classification.RuleOptions.DefaultUrgentKeywords();
            settings.ImportantKeywords = Classification.RuleOptions.DefaultImportantKeywords();
            return settings;
        }
    }
}
