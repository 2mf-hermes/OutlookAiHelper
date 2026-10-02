using System;
using System.IO;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    public static class SettingsPrivacyTests
    {
        public static int Run()
        {
            var failed = 0;
            var dir = Path.Combine(Path.GetTempPath(), "OutlookAiHelperSettings-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var settingsPath = Path.Combine(dir, "settings.json");
                var store = new JsonSettingsStore(settingsPath);

                failed += Check("defaults when missing", () =>
                {
                    var s = store.Load();
                    return s.ScanDays == 30 && s.UrgentKeywords.Count > 0 && s.AiEnabled == false;
                });

                failed += Check("save and load keywords", () =>
                {
                    var s = AppSettings.CreateDefault();
                    s.ScanDays = 14;
                    s.UrgentKeywords.Add("特急件");
                    s.VipAddresses.Add("boss@contoso.com");
                    s.Language = "en-US";
                    store.Save(s);
                    var loaded = store.Load();
                    return loaded.ScanDays == 14
                        && loaded.UrgentKeywords.Contains("特急件")
                        && loaded.VipAddresses.Contains("boss@contoso.com")
                        && loaded.Language == "en-US";
                });

                failed += Check("export copies files", () =>
                {
                    File.WriteAllText(Path.Combine(dir, "todos.json"), "{\"schemaVersion\":1,\"items\":[]}");
                    File.WriteAllText(Path.Combine(dir, "overrides.json"), "{\"schemaVersion\":1,\"items\":[]}");
                    var privacy = new PrivacyUseCase(dir);
                    var exportDir = Path.Combine(dir, "out");
                    var target = privacy.ExportAll(exportDir);
                    return File.Exists(Path.Combine(target, "settings.json"))
                        && File.Exists(Path.Combine(target, "todos.json"))
                        && File.Exists(Path.Combine(target, "overrides.json"));
                });

                failed += Check("clear removes data", () =>
                {
                    var privacy = new PrivacyUseCase(dir);
                    privacy.ClearLocalData(true, true, true);
                    return !File.Exists(Path.Combine(dir, "settings.json"))
                        && !File.Exists(Path.Combine(dir, "todos.json"))
                        && !File.Exists(Path.Combine(dir, "overrides.json"));
                });

                failed += Check("bad settings quarantines", () =>
                {
                    File.WriteAllText(settingsPath, "{ broken");
                    var loaded = new JsonSettingsStore(settingsPath).Load();
                    return loaded.ScanDays == 30;
                });

                failed += Check("null ai stays disabled", () =>
                {
                    var ai = new OutlookAiHelper.Adapters.Ai.NullAiCapability();
                    var result = ai.Suggest(new OutlookAiHelper.Core.Abstractions.AiRequest { Subject = "x" });
                    return !ai.IsEnabled && !result.Success && result.ErrorKey == "ai.disabled";
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
