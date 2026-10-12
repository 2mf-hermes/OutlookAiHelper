using System;
using System.Collections.Generic;
using System.IO;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Application;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the mail-source skeleton: what a settings file written before sources
    /// existed resolves to, how a source survives a save/load, and that nothing -- a missing
    /// source, an unknown kind, a cloud source with no usable registration, an id that could
    /// not name a file -- leaves the app without a reader.
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

        /// <summary>
        /// A sample registration id, the same in every check that needs one.
        ///
        /// A client id is public information -- a public client has no secret, which is why the
        /// device-code flow fits an app anyone can install -- and this particular one belongs to
        /// nobody, so it stands in for the real one without standing in for anything secret.
        /// </summary>
        private const string CloudId = "9a1c0e5b-2f4d-4a6e-9b31-7c5d0e2a8f14";

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

                failed += Check("an unknown kind resolves to no reader, and the local one is handed back", () =>
                {
                    var local = new FakeReader();
                    var unknown = new MailSourceProfile { Id = "unknown-1", Kind = (MailSourceKind)99 };
                    return !MailReaderFactory.IsImplemented((MailSourceKind)99)
                        && !MailReaderFactory.IsImplemented((MailSourceKind)99, CloudId)
                        && MailReaderFactory.IsImplemented(MailSourceKind.LocalOutlook)
                        && MailReaderFactory.Create((MailSourceKind)99, local) == null
                        && MailReaderFactory.Create(unknown, local, CloudId) == null
                        && ReferenceEquals(MailReaderFactory.Create(MailSourceKind.LocalOutlook, local), local)
                        && ReferenceEquals(MailReaderFactory.Create(MailSourceProfile.Local(), local, CloudId), local);
                });

                failed += Check("a cloud source with no usable registration is never built", () =>
                {
                    var cloud = new MailSourceProfile { Kind = MailSourceKind.Graph, Name = "cloud" };
                    return MailReaderFactory.Create(cloud, null, null) == null
                        && MailReaderFactory.Create(cloud, null, string.Empty) == null
                        && MailReaderFactory.Create(cloud, null, "placeholder") == null
                        && MailReaderFactory.Create(null, null, CloudId) == null
                        && !MailReaderFactory.IsImplemented(MailSourceKind.Graph, string.Empty)
                        && !MailReaderFactory.IsImplemented(MailSourceKind.Graph, "placeholder");
                });

                failed += Check("the port carries the probe and the open-item call", () =>
                {
                    IMailReader reader = new FakeReader();
                    return reader.ProbeInbox().IsUsable && !reader.TryOpenInOutlook("entry");
                });
                failed += Check("the source list always offers a way back to the local mailbox", () =>
                {
                    var configured = new List<MailSourceProfile>
                    {
                        new MailSourceProfile { Kind = MailSourceKind.Graph, Name = "cloud" }
                    };
                    var shown = MailSourceAvailability.DisplayList(configured);
                    return shown.Count == 2
                        && shown[0].Kind == MailSourceKind.LocalOutlook
                        && shown[0].Id == MailSourceProfile.LocalId
                        && shown[1].Kind == MailSourceKind.Graph;
                });

                failed += Check("a stored local source is listed once, not twice", () =>
                {
                    var configured = new List<MailSourceProfile>
                    {
                        MailSourceProfile.Local(),
                        new MailSourceProfile { Kind = MailSourceKind.Graph, Name = "cloud" }
                    };
                    var shown = MailSourceAvailability.DisplayList(configured);
                    return shown.Count == 2 && shown[0].Id == MailSourceProfile.LocalId;
                });

                failed += Check("the local source can never be removed", () =>
                {
                    var cloud = new MailSourceProfile { Kind = MailSourceKind.Graph };
                    return !MailSourceAvailability.IsRemovable(MailSourceProfile.Local())
                        && !MailSourceAvailability.IsRemovable(null)
                        && MailSourceAvailability.IsRemovable(cloud);
                });

                failed += Check("a local-only build offers nothing to add and asks nothing", () =>
                {
                    Func<MailSourceKind, bool> only = kind => kind == MailSourceKind.LocalOutlook;
                    return !MailSourceAvailability.CanOfferNew(only)
                        && !MailSourceAvailability.NeedsFirstRunChoice(only, false)
                        && MailSourceAvailability.FirstOfferableKind(only, MailSourceAvailability.CloudKinds) == null;
                });

                failed += Check("with a cloud reader the first run asks once, then never again", () =>
                {
                    Func<MailSourceKind, bool> any = kind => true;
                    return MailSourceAvailability.CanOfferNew(any)
                        && MailSourceAvailability.NeedsFirstRunChoice(any, false)
                        && !MailSourceAvailability.NeedsFirstRunChoice(any, true)
                        && MailSourceAvailability.FirstOfferableKind(any, MailSourceAvailability.CloudKinds) == MailSourceKind.Graph;
                });

                failed += Check("the local kind is never something to add", () =>
                {
                    var kinds = MailSourceAvailability.CloudKinds;
                    for (var i = 0; i < kinds.Length; i++)
                    {
                        if (kinds[i] == MailSourceKind.LocalOutlook)
                        {
                            return false;
                        }
                    }

                    return kinds.Length > 0;
                });

                failed += Check("a cloud source stays unreadable while the cloud module is hidden", () =>
                {
                    // The id below is well formed, and that is the point: while the cloud module is
                    // hidden the switch is the outer answer, so a registration this build could use
                    // still answers with the local mailbox rather than with a reader this build
                    // ships no registration for.
                    var cloud = new MailSourceProfile { Id = "cloud-1", Kind = MailSourceKind.Graph, Name = "cloud" };
                    var s = AppSettings.CreateDefault();
                    s.MailSources.Add(MailSourceProfile.Local());
                    s.MailSources.Add(cloud);
                    s.ActiveMailSourceId = cloud.Id;

                    var local = new FakeReader();
                    var runtime = new MailSourceRuntime(local, CloudId);
                    runtime.Configure(s);

                    // Reading twice is what makes this about the cache as well as the answer: the
                    // second read has to hand back the same reader, not build one per call.
                    var first = runtime.Reader;
                    return runtime.ActiveProfile.Id == MailSourceProfile.LocalId
                        && runtime.RequiresSta
                        && ReferenceEquals(first, local)
                        && ReferenceEquals(runtime.Reader, first)
                        && MailReaderFactory.Create(cloud, null, CloudId) == null;
                });

                failed += Check("switching away from a hidden cloud source hands the local reader back", () =>
                {
                    var cloud = new MailSourceProfile { Id = "cloud-1", Kind = MailSourceKind.Graph, Name = "cloud" };
                    var local = new FakeReader();
                    var runtime = new MailSourceRuntime(local, CloudId);

                    var onCloud = AppSettings.CreateDefault();
                    onCloud.MailSources.Add(MailSourceProfile.Local());
                    onCloud.MailSources.Add(cloud);
                    onCloud.ActiveMailSourceId = cloud.Id;
                    runtime.Configure(onCloud);
                    var hiddenCloud = runtime.Reader;

                    var onLocal = AppSettings.CreateDefault();
                    onLocal.MailSources.Add(MailSourceProfile.Local());
                    onLocal.ActiveMailSourceId = MailSourceProfile.LocalId;
                    runtime.Configure(onLocal);

                    // Both settings files answer with the local mailbox, so this is a switch between
                    // two different active sources that happen to resolve to the same reader: what
                    // is asked here is that the reader belongs to the source that is active now.
                    return ReferenceEquals(hiddenCloud, local)
                        && ReferenceEquals(runtime.Reader, local)
                        && runtime.ActiveProfile.Id == MailSourceProfile.LocalId;
                });

                failed += Check("a cloud source this build cannot read falls back to the local mailbox", () =>
                {
                    var cloud = new MailSourceProfile { Id = "cloud-1", Kind = MailSourceKind.Graph, Name = "cloud" };
                    var s = AppSettings.CreateDefault();
                    s.MailSources.Add(MailSourceProfile.Local());
                    s.MailSources.Add(cloud);
                    s.ActiveMailSourceId = cloud.Id;

                    var local = new FakeReader();
                    var runtime = new MailSourceRuntime(local, string.Empty);
                    runtime.Configure(s);

                    return ReferenceEquals(runtime.Reader, local)
                        && runtime.ActiveProfile.Id == MailSourceProfile.LocalId
                        && runtime.RequiresSta;
                });

                failed += Check("a source id that could not name a file gets no reader", () =>
                {
                    var escaping = new MailSourceProfile { Id = "..\\..\\secrets", Kind = MailSourceKind.Graph };
                    var slashed = new MailSourceProfile { Id = "a/b", Kind = MailSourceKind.Graph };
                    var empty = new MailSourceProfile { Id = string.Empty, Kind = MailSourceKind.Graph };
                    var tooLong = new MailSourceProfile { Id = new string('x', 65), Kind = MailSourceKind.Graph };
                    var plain = new MailSourceProfile { Id = "cloud-1", Kind = MailSourceKind.Graph };

                    // Nothing here is built while the cloud module is hidden, so the guard is asked
                    // about the ids themselves as well as through Create: a name that could point
                    // outside the secrets folder has to be refused as a name, whether or not a
                    // reader would have been built for it.
                    return MailReaderFactory.Create(escaping, null, CloudId) == null
                        && MailReaderFactory.Create(slashed, null, CloudId) == null
                        && MailReaderFactory.Create(empty, null, CloudId) == null
                        && MailReaderFactory.Create(tooLong, null, CloudId) == null
                        && MailReaderFactory.Create(plain, null, CloudId) == null
                        && JsonPaths.IsSafeSourceId("cloud-1")
                        && !JsonPaths.IsSafeSourceId(escaping.Id)
                        && !JsonPaths.IsSafeSourceId(slashed.Id)
                        && !JsonPaths.IsSafeSourceId(empty.Id)
                        && !JsonPaths.IsSafeSourceId(tooLong.Id);
                });

                failed += Check("removing the active source falls back to the local mailbox", () =>
                {
                    return MailSourceAvailability.ActiveIdAfterRemoval("cloud", "cloud") == MailSourceProfile.LocalId
                        && MailSourceAvailability.ActiveIdAfterRemoval("cloud", "other") == "cloud"
                        && MailSourceAvailability.ActiveIdAfterRemoval(null, null) == MailSourceProfile.LocalId;
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
