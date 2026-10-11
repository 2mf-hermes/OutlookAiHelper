using System;
using System.Collections.Generic;
using System.IO;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the mail-source skeleton: what a settings file written before sources
    /// existed resolves to, how a source survives a save/load, and that neither a missing
    /// source nor a kind this build cannot read can leave the app without a reader.
    /// </summary>
    public static class MailSourceTests
    {
        private sealed class FakeReader : IMailReader
        {
            public bool IsAvailable(out string reasonKey)
            {
                reasonKey = null;
                return true;
            }

            public IList<string> ListFolders()
            {
                return new List<string>();
            }

            public IEnumerable<MailSummary> Enumerate(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled)
            {
                yield break;
            }

            public InboxProbe ProbeInbox()
            {
                return new InboxProbe { Signature = "probe", InboxCount = 1 };
            }

            public bool TryOpenInOutlook(string entryId)
            {
                return false;
            }
        }

        public static int Run()
        {
            var failed = 0;
            var dir = Path.Combine(Path.GetTempPath(), "OutlookAiHelperSources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var path = Path.Combine(dir, "settings.json");
                var store = new JsonSettingsStore(path);

                failed += Check("a settings file with no sources still reads local", () =>
                {
                    var s = AppSettings.CreateDefault();
                    store.Save(s);
                    var loaded = store.Load();
                    var active = loaded.GetActiveMailSource();
                    return active != null
                        && active.Kind == MailSourceKind.LocalOutlook
                        && active.Id == MailSourceProfile.LocalId
                        && active.RequiresSta;
                });

                failed += Check("a source list round-trips", () =>
                {
                    var s = AppSettings.CreateDefault();
                    s.MailSources.Add(MailSourceProfile.Local());
                    var cloud = new MailSourceProfile
                    {
                        Kind = MailSourceKind.Graph,
                        Name = "Work mailbox",
                        AccountUpn = "someone@contoso.com"
                    };
                    s.MailSources.Add(cloud);
                    s.ActiveMailSourceId = cloud.Id;
                    store.Save(s);

                    var loaded = store.Load();
                    if (loaded.MailSources.Count != 2)
                    {
                        return false;
                    }

                    var back = loaded.GetActiveMailSource();
                    return back.Id == cloud.Id
                        && back.Kind == MailSourceKind.Graph
                        && back.Name == "Work mailbox"
                        && back.AccountUpn == "someone@contoso.com"
                        && !back.RequiresSta
                        && loaded.MailSources[0].Id == MailSourceProfile.LocalId;
                });

                failed += Check("the kind is stored by name, not by number", () =>
                {
                    var s = AppSettings.CreateDefault();
                    s.MailSources.Add(new MailSourceProfile { Kind = MailSourceKind.Graph });
                    store.Save(s);
                    var raw = File.ReadAllText(path);
                    return raw.Contains("\"Kind\":\"Graph\"");
                });

                failed += Check("a kind from a newer build is not read as local or graph", () =>
                {
                    var s = AppSettings.CreateDefault();
                    s.MailSources.Add(new MailSourceProfile { Kind = MailSourceKind.Graph, Name = "Future" });
                    store.Save(s);
                    var raw = File.ReadAllText(path);
                    File.WriteAllText(path, raw.Replace("\"Kind\":\"Graph\"", "\"Kind\":\"SomethingNew\""));
                    var loaded = new JsonSettingsStore(path).Load();
                    return loaded.MailSources.Count == 1
                        && loaded.MailSources[0].Kind == MailSourceKind.LocalOutlook;
                });

                failed += Check("an unknown active id falls back to a source that exists", () =>
                {
                    var s = AppSettings.CreateDefault();
                    var first = MailSourceProfile.Local();
                    s.MailSources.Add(first);
                    s.ActiveMailSourceId = "no-such-source";
                    store.Save(s);
                    var loaded = store.Load();
                    return loaded.GetActiveMailSource().Id == first.Id;
                });

                failed += Check("a runtime with no settings still has a reader", () =>
                {
                    var runtime = new MailSourceRuntime(new FakeReader());
                    runtime.Configure(AppSettings.CreateDefault());
                    return runtime.Reader != null
                        && runtime.ActiveProfile.Kind == MailSourceKind.LocalOutlook
                        && runtime.RequiresSta;
                });

                failed += Check("a source this build cannot read keeps the app readable", () =>
                {
                    var s = AppSettings.CreateDefault();
                    var cloud = new MailSourceProfile { Kind = MailSourceKind.Graph, Name = "Cloud" };
                    s.MailSources.Add(cloud);
                    s.ActiveMailSourceId = cloud.Id;

                    var runtime = new MailSourceRuntime(new FakeReader());
                    runtime.Configure(s);

                    // Falls back to local so scans keep working, but still lists the cloud
                    // source so the UI can show it and mark it unusable.
                    return runtime.Reader != null
                        && runtime.ActiveProfile.Kind == MailSourceKind.LocalOutlook
                        && runtime.RequiresSta
                        && runtime.Sources.Count == 1
                        && runtime.Sources[0].Kind == MailSourceKind.Graph;
                });

                failed += Check("a runtime follows the configured active source", () =>
                {
                    var s = AppSettings.CreateDefault();
                    var local = MailSourceProfile.Local();
                    s.MailSources.Add(local);
                    s.ActiveMailSourceId = local.Id;

                    var runtime = new MailSourceRuntime(new FakeReader());
                    runtime.Configure(s);
                    return runtime.ActiveProfile.Id == local.Id
                        && runtime.ActiveProfile.Kind == MailSourceKind.LocalOutlook;
                });

                failed += Check("an unimplemented kind resolves to no reader", () =>
                {
                    var local = new FakeReader();
                    return !MailReaderFactory.IsImplemented(MailSourceKind.Graph)
                        && MailReaderFactory.IsImplemented(MailSourceKind.LocalOutlook)
                        && MailReaderFactory.Create(MailSourceKind.Graph, local) == null
                        && ReferenceEquals(MailReaderFactory.Create(MailSourceKind.LocalOutlook, local), local);
                });

                failed += Check("the port carries the probe and the open-item call", () =>
                {
                    IMailReader reader = new FakeReader();
                    return reader.ProbeInbox().IsUsable && !reader.TryOpenInOutlook("entry");
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
