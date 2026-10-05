using System;
using System.IO;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the automatic-refresh plumbing that does not need Outlook: the probe's
    /// signature rules, and the settings that decide whether the probe runs at all.
    /// The COM probe itself is exercised by hand, since it needs a live mailbox.
    /// </summary>
    public static class AutoRefreshTests
    {
        public static int Run()
        {
            var failed = 0;
            var dir = Path.Combine(Path.GetTempPath(), "OutlookAiHelperRefresh-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var settingsPath = Path.Combine(dir, "settings.json");
                var store = new JsonSettingsStore(settingsPath);

                failed += Check("probe signature tracks the inbox head", () =>
                {
                    var first = InboxProbe.MakeFragment("AAA", new DateTime(2026, 9, 1, 1, 0, 0, DateTimeKind.Utc));
                    var sameAgain = InboxProbe.MakeFragment("AAA", new DateTime(2026, 9, 1, 1, 0, 0, DateTimeKind.Utc));
                    var newerMail = InboxProbe.MakeFragment("BBB", new DateTime(2026, 9, 1, 1, 0, 0, DateTimeKind.Utc));
                    var newerTime = InboxProbe.MakeFragment("AAA", new DateTime(2026, 9, 2, 1, 0, 0, DateTimeKind.Utc));
                    return first == sameAgain && first != newerMail && first != newerTime;
                });

                failed += Check("probe signature ignores the timezone it is written in", () =>
                {
                    var utc = InboxProbe.MakeFragment("AAA", new DateTime(2026, 9, 1, 1, 0, 0, DateTimeKind.Utc));
                    var local = InboxProbe.MakeFragment("AAA", new DateTime(2026, 9, 1, 1, 0, 0, DateTimeKind.Utc).ToLocalTime());
                    return utc == local;
                });

                failed += Check("an unreachable inbox is unknown, not unchanged", () =>
                {
                    var empty = new InboxProbe { Signature = null, InboxCount = 0 };
                    var found = new InboxProbe { Signature = "AAA@1", InboxCount = 1 };
                    return !empty.IsUsable && found.IsUsable;
                });

                failed += Check("probe interval defaults when the field is absent", () =>
                {
                    // An older settings file simply has no such field. Reading it as 0
                    // would silently switch the feature off for everyone who upgrades.
                    File.WriteAllText(settingsPath, "{\"SchemaVersion\":1,\"Language\":\"zh-TW\",\"ScanDays\":30}");
                    var loaded = new JsonSettingsStore(settingsPath).Load();
                    return loaded.AutoRefreshMinutes == AppSettings.DefaultAutoRefreshMinutes;
                });

                failed += Check("an explicit off survives a round trip", () =>
                {
                    var settings = AppSettings.CreateDefault();
                    settings.AutoRefreshMinutes = 0;
                    store.Save(settings);
                    return new JsonSettingsStore(settingsPath).Load().AutoRefreshMinutes == 0;
                });

                failed += Check("an absurd interval falls back to the default", () =>
                {
                    File.WriteAllText(settingsPath, "{\"SchemaVersion\":1,\"AutoRefreshMinutes\":9999}");
                    var loaded = new JsonSettingsStore(settingsPath).Load();
                    var saved = AppSettings.CreateDefault();
                    saved.AutoRefreshMinutes = -5;
                    store.Save(saved);
                    return loaded.AutoRefreshMinutes == AppSettings.DefaultAutoRefreshMinutes
                        && new JsonSettingsStore(settingsPath).Load().AutoRefreshMinutes == AppSettings.DefaultAutoRefreshMinutes;
                });

                failed += Check("the update check preference persists", () =>
                {
                    // Regression: this setting was written to the model but never to the
                    // file, so "check on startup" quietly reset to off on every restart.
                    var settings = AppSettings.CreateDefault();
                    settings.CheckForUpdatesOnStartup = true;
                    store.Save(settings);
                    return new JsonSettingsStore(settingsPath).Load().CheckForUpdatesOnStartup;
                });

                failed += Check("the update check stays off by default", () =>
                {
                    File.WriteAllText(settingsPath, "{\"SchemaVersion\":1,\"ScanDays\":30}");
                    return !new JsonSettingsStore(settingsPath).Load().CheckForUpdatesOnStartup;
                });
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch (Exception) { }
            }

            return failed;
        }

        private static int Check(string name, Func<bool> body)
        {
            try
            {
                if (body())
                {
                    Console.WriteLine("PASS " + name);
                    return 0;
                }

                Console.WriteLine("FAIL " + name);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + " :: " + ex.Message);
                return 1;
            }
        }
    }
}
