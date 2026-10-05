using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Outlook
{
    /// <summary>
    /// Read-only Outlook COM reader. Call only from an STA thread (see ComSta).
    /// Caps volume to avoid memory spikes on huge mailboxes.
    /// </summary>
    public sealed class OutlookComMailReader : IMailReader
    {
        private const int MaxItems = 2000;

        /// <summary>Outlook OlDefaultFolders.olFolderInbox.</summary>
        private const int OlFolderInbox = 6;

        /// <summary>
        /// How many items at the head of an Inbox the auto-refresh probe looks at when it
        /// picks the newest one. Small on purpose: the probe must stay far cheaper than a
        /// scan, and Outlook hands a default Inbox back newest-first.
        /// </summary>
        private const int ProbeHeadSample = 10;

        private object _application;
        private bool _attached;

        public bool IsAvailable(out string reasonKey)
        {
            ResetSession();
            object app;
            if (!OutlookAvailability.TryGetRunningApplication(out app, out reasonKey))
            {
                FileLogger.Info("IsAvailable=false " + reasonKey);
                return false;
            }

            _application = app;
            _attached = true;
            reasonKey = null;
            return true;
        }

        /// <summary>Open a mail item in Outlook by EntryID (read-only display).</summary>
        public bool OpenItemInOutlook(string entryId)
        {
            if (string.IsNullOrEmpty(entryId))
            {
                return false;
            }

            string reason;
            if (!IsAvailable(out reason))
            {
                return false;
            }

            object ns = null;
            object item = null;
            try
            {
                ns = Invoke(_application, "GetNamespace", "MAPI");
                item = Invoke(ns, "GetItemFromID", entryId);
                if (item == null)
                {
                    return false;
                }

                Invoke(item, "Display");
                return true;
            }
            catch (Exception ex)
            {
                FileLogger.Error("OpenItemInOutlook", ex);
                return false;
            }
            finally
            {
                OutlookAvailability.Release(item);
                OutlookAvailability.Release(ns);
            }
        }

        public void ResetSession()
        {
            if (_application != null)
            {
                OutlookAvailability.Release(_application);
            }

            _application = null;
            _attached = false;
        }

        public void DisposeAttachment()
        {
            ResetSession();
        }

        public IList<string> ListFolders()
        {
            var names = new List<string> { "Inbox" };
            string reason;
            if (!IsAvailable(out reason))
            {
                return names;
            }

            object ns = null;
            try
            {
                ns = Invoke(_application, "GetNamespace", "MAPI");
                object folders = null;
                try
                {
                    folders = Invoke(ns, "Folders");
                    var count = ToInt(Invoke(folders, "Count"));
                    for (var i = 1; i <= Math.Min(count, 20); i++)
                    {
                        object folder = null;
                        try
                        {
                            folder = Invoke(folders, "Item", i);
                            var name = ToText(Invoke(folder, "Name"));
                            if (!string.IsNullOrEmpty(name))
                            {
                                names.Add(name);
                            }
                        }
                        finally
                        {
                            OutlookAvailability.Release(folder);
                        }
                    }
                }
                finally
                {
                    OutlookAvailability.Release(folders);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("ListFolders failed", ex);
            }
            finally
            {
                OutlookAvailability.Release(ns);
            }

            return names;
        }

        public IEnumerable<MailSummary> Enumerate(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled)
        {
            var results = new List<MailSummary>();
            string reason;
            if (!IsAvailable(out reason))
            {
                throw new InvalidOperationException(reason ?? "error.outlook.not_detected");
            }

            object ns = null;
            try
            {
                ns = Invoke(_application, "GetNamespace", "MAPI");
                var cutoff = DateTime.Now.AddDays(-Math.Max(1, scope == null ? 30 : scope.Days));
                // Deleted and junk mail are out of scope: collect each store's default
                // Deleted Items / Junk E-mail EntryIDs so the walk can prune them (and
                // their subfolders) even when the folders were renamed.
                var excludedFolderIds = new List<string>();
                AddDefaultFolderId(ns, ScanFolderFilter.OlFolderDeletedItems, excludedFolderIds);
                AddDefaultFolderId(ns, ScanFolderFilter.OlFolderJunk, excludedFolderIds);
                var folders = new List<object>();
                CollectFolders(ns, folders, excludedFolderIds);
                if (folders.Count == 0)
                {
                    AddFallbackFolders(ns, folders);
                }
                FileLogger.Info("Scan folders=" + folders.Count + " days=" + (scope != null ? scope.Days : 0)
                    + " excludedFolders=" + excludedFolderIds.Count + " (deleted items + junk e-mail)");

                var mapped = 0;
                foreach (var folder in folders)
                {
                    if (isCancelled != null && isCancelled())
                    {
                        break;
                    }

                    if (mapped >= MaxItems)
                    {
                        break;
                    }

                    try
                    {
                        ScanFolder(folder, cutoff, results, ref mapped, progress, isCancelled, excludedFolderIds);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Error("ScanFolder failed", ex);
                    }
                    finally
                    {
                        OutlookAvailability.Release(folder);
                    }
                }

                FileLogger.Info("Scan done mapped=" + results.Count);
            }
            catch (Exception ex)
            {
                FileLogger.Error("Enumerate failed", ex);
                throw;
            }
            finally
            {
                OutlookAvailability.Release(ns);
                ResetSession();
            }

            return results;
        }

        /// <summary>
        /// Walk every MAPI store (primary + PST/archive) so rules-filed mail is included.
        /// </summary>
        private void AddFallbackFolders(object ns, List<object> folders)
        {
            try
            {
                var inbox = Invoke(ns, "GetDefaultFolder", OlFolderInbox);
                if (inbox != null)
                {
                    folders.Add(inbox);
                }
                var top = Invoke(ns, "Folders");
                try
                {
                    var n = ToInt(Invoke(top, "Count"));
                    for (var i = 1; i <= n && folders.Count < 80; i++)
                    {
                        object f = null;
                        try
                        {
                            f = Invoke(top, "Item", i);
                            if (f != null) folders.Add(f);
                        }
                        catch (Exception)
                        {
                            OutlookAvailability.Release(f);
                        }
                    }
                }
                finally
                {
                    OutlookAvailability.Release(top);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("AddFallbackFolders", ex);
            }
        }

        private void CollectFolders(object ns, List<object> folders, ICollection<string> excludedFolderIds)
        {
            try
            {
                var stores = Invoke(ns, "Folders");
                var storeCount = ToInt(Invoke(stores, "Count"));
                FileLogger.Info("Mail stores=" + storeCount);
                for (var i = 1; i <= storeCount && folders.Count < 300; i++)
                {
                    object store = null;
                    try
                    {
                        store = Invoke(stores, "Item", i);
                        var name = ToText(Invoke(store, "Name")) ?? ("store" + i);
                        FileLogger.Info("Walk store " + name);
                        // Archive PSTs keep their own Deleted Items and Junk E-mail folders.
                        AddDefaultFolderId(store, ScanFolderFilter.OlFolderDeletedItems, excludedFolderIds);
                        AddDefaultFolderId(store, ScanFolderFilter.OlFolderJunk, excludedFolderIds);
                        AddFolderRecursive(store, folders, 0, name, excludedFolderIds);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Error("Walk store failed", ex);
                    }
                    finally
                    {
                        // keep child folder refs; release only the enumeration handle
                        if (store != null && !ContainsSame(folders, store))
                        {
                            OutlookAvailability.Release(store);
                        }
                    }
                }

                OutlookAvailability.Release(stores);
            }
            catch (Exception ex)
            {
                FileLogger.Error("CollectFolders failed", ex);
            }

            // Fallback: default Inbox + its parent tree if nothing collected
            if (folders.Count == 0)
            {
                try
                {
                    var inbox = Invoke(ns, "GetDefaultFolder", OlFolderInbox);
                    if (inbox != null)
                    {
                        folders.Add(inbox);
                        FileLogger.Info("Fallback Inbox only");
                    }
                }
                catch (Exception ex)
                {
                    FileLogger.Error("Fallback inbox failed", ex);
                }
            }
        }

        private static bool ContainsSame(List<object> list, object target)
        {
            foreach (var o in list)
            {
                if (ReferenceEquals(o, target))
                {
                    return true;
                }
            }
            return false;
        }

        private void AddFolderRecursive(object folder, List<object> folders, int depth, string path, ICollection<string> excludedFolderIds)
        {
            if (folder == null || depth > 8 || folders.Count >= 300)
            {
                return;
            }

            var name = ToText(Invoke(folder, "Name")) ?? "";

            // Deleted and junk mail are not scanned: skip the folder and its whole subtree.
            if (IsExcludedFolder(folder, name, excludedFolderIds))
            {
                FileLogger.Info("Folder- (deleted, skipped) " + path);
                return;
            }

            var defaultClass = string.Empty;
            try
            {
                defaultClass = ToText(Invoke(folder, "DefaultMessageClass")) ?? string.Empty;
            }
            catch (Exception)
            {
            }

            // Skip contact/calendar/task folders; keep mail (IPM.Note*) and unknown with items
            var isMailish = defaultClass.Length == 0
                || defaultClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)
                || defaultClass.StartsWith("IPM.Schedule", StringComparison.OrdinalIgnoreCase);

            if (isMailish)
            {
                folders.Add(folder);
                FileLogger.Info("Folder+ " + path);
            }

            object children = null;
            try
            {
                children = Invoke(folder, "Folders");
                var count = ToInt(Invoke(children, "Count"));
                for (var i = 1; i <= count && folders.Count < 300; i++)
                {
                    object child = null;
                    try
                    {
                        child = Invoke(children, "Item", i);
                        var childName = ToText(Invoke(child, "Name")) ?? ("f" + i);
                        AddFolderRecursive(child, folders, depth + 1, path + "/" + childName, excludedFolderIds);
                        child = null; // ownership transferred to list or released in child
                    }
                    finally
                    {
                        if (child != null)
                        {
                            OutlookAvailability.Release(child);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("Folder children failed " + path, ex);
            }
            finally
            {
                OutlookAvailability.Release(children);
            }
        }

        private void ScanFolder(object folder, DateTime cutoff, List<MailSummary> results, ref int mapped, Action<int, int> progress, Func<bool> isCancelled, ICollection<string> excludedFolderIds)
        {
            // Last line of defence: fallback folder collection adds store roots directly,
            // bypassing AddFolderRecursive.
            if (IsExcludedFolder(folder, excludedFolderIds))
            {
                FileLogger.Info("Scan skipped deleted folder");
                return;
            }

            object items = null;
            object restricted = null;
            try
            {
                items = Invoke(folder, "Items");
                Invoke(items, "Sort", "[ReceivedTime]", true);
                var filter = "@SQL=\"urn:schemas:httpmail:datereceived\" >= '" + cutoff.ToString("g") + "'";
                restricted = Invoke(items, "Restrict", filter);
                var list = restricted ?? items;
                var count = ToInt(Invoke(list, "Count"));

                for (var i = 1; i <= count && mapped < MaxItems; i++)
                {
                    if (isCancelled != null && isCancelled())
                    {
                        return;
                    }

                    object item = null;
                    try
                    {
                        item = Invoke(list, "Item", i);
                        var summary = MapItem(item, folder);
                        if (summary != null)
                        {
                            results.Add(summary);
                            mapped++;
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        OutlookAvailability.Release(item);
                    }

                    if (progress != null && (i % 25 == 0))
                    {
                        progress(results.Count, MaxItems);
                    }
                }
            }
            finally
            {
                OutlookAvailability.Release(restricted);
                if (restricted != null || items != null)
                {
                    OutlookAvailability.Release(items);
                }
            }
        }

        public bool TryOpenInOutlook(string entryId)
        {
            if (string.IsNullOrEmpty(entryId) || entryId.StartsWith("fallback:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string reason;
            if (!IsAvailable(out reason))
            {
                return false;
            }

            object ns = null;
            object item = null;
            try
            {
                ns = Invoke(_application, "GetNamespace", "MAPI");
                item = Invoke(ns, "GetItemFromID", entryId);
                if (item == null)
                {
                    return false;
                }

                Invoke(item, "Display", false);
                FileLogger.Info("Opened mail in Outlook");
                return true;
            }
            catch (Exception ex)
            {
                FileLogger.Error("TryOpenInOutlook failed", ex);
                return false;
            }
            finally
            {
                OutlookAvailability.Release(item);
                OutlookAvailability.Release(ns);
            }
        }

        private static MailSummary MapItem(object item, object folder)
        {
            if (item == null)
            {
                return null;
            }

            try
            {
                var entryId = ToText(Invoke(item, "EntryID"));
                var subject = ToText(Invoke(item, "Subject")) ?? string.Empty;
                var senderName = ToText(Invoke(item, "SenderName")) ?? string.Empty;
                var senderEmail = ToText(Invoke(item, "SenderEmailAddress")) ?? string.Empty;
                var received = ToDate(Invoke(item, "ReceivedTime"), DateTime.Now);
                var isRead = !ToBool(Invoke(item, "UnRead"));
                var hasFlag = ToInt(Invoke(item, "FlagStatus")) != 0;
                var importance = ToInt(Invoke(item, "Importance"));
                var objectClass = ToInt(Invoke(item, "Class"));

                return new MailSummary
                {
                    EntryId = entryId ?? ("fallback:" + received.ToString("o") + ":" + subject.GetHashCode().ToString("x")),
                    FolderPath = ToText(Invoke(folder, "Name")) ?? "Inbox",
                    Subject = subject,
                    FromName = senderName,
                    FromAddress = NormalizeAddress(senderEmail),
                    ReceivedOn = received,
                    IsRead = isRead,
                    HasFlag = hasFlag,
                    IsMeetingRequest = objectClass == 53 || objectClass == 54,
                    Importance = importance
                };
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Resolve an Outlook default folder (e.g. olFolderDeletedItems) on a store or
        /// namespace and remember its EntryID, so a renamed folder is still recognised.
        /// </summary>
        private static void AddDefaultFolderId(object owner, int folderType, ICollection<string> ids)
        {
            if (owner == null || ids == null)
            {
                return;
            }

            object folder = null;
            try
            {
                folder = Invoke(owner, "GetDefaultFolder", folderType);
                var id = folder == null ? null : ToText(Invoke(folder, "EntryID"));
                if (!string.IsNullOrEmpty(id) && !ids.Contains(id))
                {
                    ids.Add(id);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Info("Default folder " + folderType + " unavailable: " + ex.Message);
            }
            finally
            {
                OutlookAvailability.Release(folder);
            }
        }

        private static bool IsExcludedFolder(object folder, ICollection<string> excludedFolderIds)
        {
            if (folder == null)
            {
                return false;
            }

            string name = null;
            try
            {
                name = ToText(Invoke(folder, "Name"));
            }
            catch (Exception)
            {
            }

            return IsExcludedFolder(folder, name, excludedFolderIds);
        }

        private static bool IsExcludedFolder(object folder, string name, ICollection<string> excludedFolderIds)
        {
            if (folder == null)
            {
                return false;
            }

            if (ScanFolderFilter.IsExcludedFolderName(name))
            {
                return true;
            }

            if (excludedFolderIds == null || excludedFolderIds.Count == 0)
            {
                return false;
            }

            string entryId = null;
            try
            {
                entryId = ToText(Invoke(folder, "EntryID"));
            }
            catch (Exception)
            {
            }

            return ScanFolderFilter.IsExcluded(name, entryId, excludedFolderIds);
        }

        /// <summary>
        /// Cheap "is there anything new?" probe for the auto-refresh timer: reads the head
        /// of each store's Inbox instead of walking every folder, so it costs a fraction of
        /// a scan. Returns null when Outlook is not reachable, which means "unknown" — the
        /// caller must skip the tick rather than conclude nothing changed.
        ///
        /// Deliberately read-only: the folder view is never sorted or re-filtered (that
        /// would reorder what the user sees in Outlook), so the newest item is taken as the
        /// newest ReceivedTime among the first few items the view already hands back.
        /// Call only from an STA thread (see ComSta).
        /// </summary>
        public InboxProbe ProbeInbox()
        {
            string reason;
            if (!IsAvailable(out reason))
            {
                return null;
            }

            var fragments = new List<string>();
            var unread = 0;
            var inboxes = 0;
            object ns = null;
            object stores = null;
            try
            {
                ns = Invoke(_application, "GetNamespace", "MAPI");
                stores = Invoke(ns, "Folders");
                var storeCount = ToInt(Invoke(stores, "Count"));
                for (var i = 1; i <= storeCount; i++)
                {
                    object store = null;
                    object inbox = null;
                    try
                    {
                        store = Invoke(stores, "Item", i);
                        inbox = Invoke(store, "GetDefaultFolder", OlFolderInbox);
                        if (inbox == null)
                        {
                            continue;
                        }

                        inboxes++;
                        var unreadHere = Math.Max(0, InboxUnreadCount(inbox));
                        unread += unreadHere;
                        fragments.Add(InboxHeadFragment(inbox, unreadHere));
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Info("Inbox probe skipped a store: " + ex.Message);
                    }
                    finally
                    {
                        OutlookAvailability.Release(inbox);
                        OutlookAvailability.Release(store);
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("ProbeInbox", ex);
                return null;
            }
            finally
            {
                OutlookAvailability.Release(stores);
                OutlookAvailability.Release(ns);
            }

            var probe = new InboxProbe
            {
                Signature = string.Join(";", fragments.ToArray()),
                UnreadCount = unread,
                InboxCount = inboxes
            };
            FileLogger.Info("Inbox probe inboxes=" + inboxes + " unread=" + unread + " usable=" + probe.IsUsable);
            return probe;
        }

        /// <summary>
        /// Signature fragment for one Inbox: its newest item plus its unread count.
        /// "empty" is a stable value, so an empty Inbox must not look like a change on
        /// every tick (that would turn the probe into a rescan loop).
        /// </summary>
        private static string InboxHeadFragment(object inbox, int unreadCount)
        {
            var unreadPart = Math.Max(0, unreadCount).ToString();
            object items = null;
            try
            {
                items = Invoke(inbox, "Items");
                var count = ToInt(Invoke(items, "Count"));
                var sample = count < ProbeHeadSample ? count : ProbeHeadSample;
                var bestId = string.Empty;
                var bestReceived = DateTime.MinValue;
                for (var i = 1; i <= sample; i++)
                {
                    object item = null;
                    try
                    {
                        item = Invoke(items, "Item", i);
                        var received = ToDate(Invoke(item, "ReceivedTime"), DateTime.MinValue);
                        if (bestReceived == DateTime.MinValue || received > bestReceived)
                        {
                            bestReceived = received;
                            bestId = ToText(Invoke(item, "EntryID")) ?? string.Empty;
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        OutlookAvailability.Release(item);
                    }
                }

                var head = string.IsNullOrEmpty(bestId) ? "empty" : InboxProbe.MakeFragment(bestId, bestReceived);
                return head + "#" + unreadPart;
            }
            catch (Exception ex)
            {
                FileLogger.Info("Inbox head unreadable: " + ex.Message);
                return "unknown#" + unreadPart;
            }
            finally
            {
                OutlookAvailability.Release(items);
            }
        }

        /// <summary>Unread items in a folder. Outlook spells the property UnReadItemCount.</summary>
        private static int InboxUnreadCount(object folder)
        {
            try
            {
                return ToInt(Invoke(folder, "UnReadItemCount"));
            }
            catch (Exception)
            {
            }

            try
            {
                return ToInt(Invoke(folder, "UnreadItemCount"));
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static string NormalizeAddress(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            if (value.StartsWith("EX:/", StringComparison.OrdinalIgnoreCase) || value.StartsWith("/o=", StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }

            var start = value.IndexOf('<');
            var end = value.IndexOf('>');
            if (start >= 0 && end > start)
            {
                return value.Substring(start + 1, end - start - 1).Trim();
            }

            return value.Trim();
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                null,
                target,
                args);
        }

        private static string ToText(object value)
        {
            return value == null ? null : Convert.ToString(value);
        }

        private static int ToInt(object value)
        {
            try
            {
                return value == null ? 0 : Convert.ToInt32(value);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool ToBool(object value)
        {
            try
            {
                return value != null && Convert.ToBoolean(value);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static DateTime ToDate(object value, DateTime fallback)
        {
            if (value is DateTime)
            {
                return (DateTime)value;
            }

            DateTime parsed;
            if (value != null && DateTime.TryParse(value.ToString(), out parsed))
            {
                return parsed;
            }

            return fallback;
        }
    }
}
