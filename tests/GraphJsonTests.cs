using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;
using OutlookAiHelper.Adapters.Mail;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// What Graph's json turns into before anything above it looks at a mail. The desktop reader
    /// and the cloud reader have to describe the same mail the same way -- same fields, same
    /// meaning for "flagged" and for "is a meeting" -- or the mail list, the sort and the filters
    /// would behave differently depending on where a mailbox was read from.
    ///
    /// The payloads here are hand-written copies of what Graph documents for these fields. They
    /// pin the mapping, not the service: whether a real tenant sends exactly these values is a
    /// question only a signed-in account can answer, and it is recorded as still open.
    /// </summary>
    public static class GraphJsonTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("a full message fills every field the app reads", () =>
            {
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-1",
                    ",\"subject\":\"Quarterly report\""
                    + ",\"from\":{\"emailAddress\":{\"name\":\"Ada Lovelace\",\"address\":\"ada@example.com\"}}"
                    + ",\"receivedDateTime\":\"2024-05-06T07:08:09Z\""
                    + ",\"isRead\":false"
                    + ",\"flag\":{\"flagStatus\":\"flagged\"}"
                    + ",\"importance\":\"high\""
                    + ",\"meetingMessageType\":\"meetingRequest\"")), "Inbox");

                if (rows == null || rows.Count != 1)
                {
                    return false;
                }

                var row = rows[0];
                return row.EntryId == "AAMkAD-1"
                    && row.FolderPath == "Inbox"
                    && row.Subject == "Quarterly report"
                    && row.FromName == "Ada Lovelace"
                    && row.FromAddress == "ada@example.com"
                    && row.ReceivedOn.ToUniversalTime() == new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc)
                    && row.ReceivedOn.Kind == DateTimeKind.Local
                    && !row.IsRead
                    && row.HasFlag
                    && row.IsMeetingRequest
                    && row.Importance == 2;
            });

            failed += Check("the folder a message was read from is what the scan asked for", () =>
            {
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-2", "")), "Archive");
                return rows != null && rows.Count == 1 && rows[0].FolderPath == "Archive";
            });

            failed += Check("a row that does not say whether it was read counts as unread", () =>
            {
                var before = DateTime.Now.AddMinutes(-1);
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-3", "")), "Inbox");
                if (rows == null || rows.Count != 1)
                {
                    return false;
                }

                var row = rows[0];
                return row.Subject == string.Empty
                    && row.FromName == string.Empty
                    && row.FromAddress == string.Empty
                    && !row.IsRead
                    && !row.HasFlag
                    && !row.IsMeetingRequest
                    && row.Importance == 1
                    && row.ReceivedOn >= before;
            });

            failed += Check("a from without a name or an address is not a crash", () =>
            {
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-4", ",\"from\":{}")), "Inbox");
                return rows != null && rows.Count == 1
                    && rows[0].FromName == string.Empty
                    && rows[0].FromAddress == string.Empty;
            });

            failed += Check("an offset received time is kept as the instant it names", () =>
            {
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-5",
                    ",\"receivedDateTime\":\"2024-05-06T15:08:09+08:00\"")), "Inbox");
                return rows != null && rows.Count == 1
                    && rows[0].ReceivedOn.ToUniversalTime() == new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);
            });

            failed += Check("a flag that cannot be read is not a flagged mail", () =>
            {
                var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-6",
                    ",\"receivedDateTime\":\"2024-05-06T07:08:09Z\""
                    + ",\"flag\":{\"flagStatus\":\"garbage\"}")), "Inbox");
                return rows != null && rows.Count == 1 && !rows[0].HasFlag;
            });

            failed += Check("a completed flag counts as flagged, the way Outlook counts it", () =>
                Flagged("complete") && Flagged("flagged") && Flagged("FLAGGED")
                && !Flagged("notFlagged") && !Flagged(null));

            failed += Check("importance maps onto Outlook's numbers", () =>
                Importance("low") == 0
                && Importance("normal") == 1
                && Importance("high") == 2
                && Importance("nonsense") == 1
                && Importance(null) == 1);

            failed += Check("a meeting and its cancellation count, a reply to one does not", () =>
                Meeting("meetingRequest") && Meeting("meetingCancelled") && Meeting("MEETINGREQUEST")
                && !Meeting("meetingAccepted") && !Meeting("meetingTentativelyAccepted")
                && !Meeting("meetingDeclined") && !Meeting("meetingForwardNotification")
                && !Meeting("none") && !Meeting(null));

            failed += Check("a message with no id is dropped rather than shown unopenable", () =>
            {
                var rows = GraphJson.ParseMessages(Page(
                    "{\"subject\":\"no id\"},"
                    + Message("AAMkAD-7", "")
                    + ",null"), "Inbox");
                return rows != null && rows.Count == 1 && rows[0].EntryId == "AAMkAD-7";
            });

            failed += Check("an empty folder is an empty list rather than an unreadable answer", () =>
            {
                var empty = GraphJson.ParseMessages(Page(""), "Inbox");
                return empty != null && empty.Count == 0;
            });

            failed += Check("anything that is not a json object is an unreadable answer", () =>
                GraphJson.ParseMessages(null, "Inbox") == null
                && GraphJson.ParseMessages(string.Empty, "Inbox") == null
                && GraphJson.ParseMessages("   ", "Inbox") == null
                && GraphJson.ParseMessages("not json at all", "Inbox") == null
                && GraphJson.ParseMessages("{", "Inbox") == null
                && GraphJson.ParseMessages("[1,2,3]", "Inbox") == null
                && GraphJson.ParseMessages("\"a string\"", "Inbox") == null
                && GraphJson.ParseMessages("<html><body>Sign in</body></html>", "Inbox") == null
                && GraphJson.ParseFolders("<html><body>Sign in</body></html>") == null
                && GraphJson.ParseFolders("[1,2,3]") == null);

            failed += Check("a page whose value is missing or null is empty, not unreadable", () =>
            {
                var missing = GraphJson.ParseMessages("{\"@odata.context\":\"https://graph.microsoft.com/v1.0/$metadata#users\"}", "Inbox");
                var nulled = GraphJson.ParseMessages("{\"value\":null}", "Inbox");
                return missing != null && missing.Count == 0
                    && nulled != null && nulled.Count == 0;
            });

            failed += Check("the folder list puts the inbox first and drops nameless folders", () =>
            {
                var folders = GraphJson.ParseFolders("{\"value\":["
                    + "{\"id\":\"f-archive\",\"displayName\":\"Archive\"},"
                    + "{\"id\":\"f-inbox\",\"displayName\":\"Posteingang\",\"wellKnownName\":\"inbox\"},"
                    + "{\"displayName\":\"No id\"},"
                    + "{\"id\":\"f-noname\"},"
                    + "{\"id\":\"f-drafts\",\"displayName\":\"Drafts\"}"
                    + "]}");

                if (folders == null || folders.Count != 3)
                {
                    return false;
                }

                var names = GraphJson.FolderNames(folders);
                return names.Count == 3
                    && names[0] == "Posteingang"
                    && names[1] == "Archive"
                    && names[2] == "Drafts"
                    && GraphJson.ParseFolders("broken") == null
                    && GraphJson.ParseFolders(string.Empty) == null
                    && GraphJson.FolderNames(null).Count == 0;
            });

            failed += Check("Inbox resolves by well-known name, any other folder by name", () =>
            {
                var folders = GraphJson.ParseFolders("{\"value\":["
                    + "{\"id\":\"f-archive\",\"displayName\":\"Archive\"},"
                    + "{\"id\":\"f-inbox\",\"displayName\":\"Posteingang\",\"wellKnownName\":\"inbox\"}"
                    + "]}");

                return GraphJson.ResolveFolderId(folders, "Archive") == "f-archive"
                    && GraphJson.ResolveFolderId(folders, "archive") == "f-archive"
                    && GraphJson.ResolveFolderId(folders, " Inbox ") == "f-inbox"
                    && GraphJson.ResolveFolderId(folders, "Posteingang") == "f-inbox"
                    && GraphJson.ResolveFolderId(folders, "Nothing") == null
                    && GraphJson.ResolveFolderId(folders, null) == null
                    && GraphJson.ResolveFolderId(null, "Inbox") == null;
            });

            failed += Check("a Graph error gives up its code, anything else gives up nothing", () =>
                GraphJson.ParseErrorCode("{\"error\":{\"code\":\"invalid_grant\",\"message\":\"AADSTS70000\"}}") == "invalid_grant"
                && GraphJson.ParseErrorCode("{\"error\":{\"message\":\"no code\"}}") == null
                && GraphJson.ParseErrorCode("{\"error\":\"flat\"}") == null
                && GraphJson.ParseErrorCode("{}") == null
                && GraphJson.ParseErrorCode("not json") == null);

            failed += Check("every field the parser reads is a field the request asks for", () =>
            {
                // Asked for: the parser must read nothing that was not requested.
                foreach (var field in GraphJson.MessageReadFields)
                {
                    if (!Has(GraphApp.SelectFields, field) && !Has(NestedFields, field))
                    {
                        return false;
                    }
                }

                // Requested: nothing may be asked for that the parser ignores.
                foreach (var field in GraphApp.SelectFields)
                {
                    if (!Has(GraphJson.MessageReadFields, field))
                    {
                        return false;
                    }
                }

                foreach (var field in GraphApp.FolderSelectFields)
                {
                    if (!Has(GraphJson.FolderReadFields, field) || GraphApp.FolderListUrl.IndexOf(field, StringComparison.Ordinal) < 0)
                    {
                        return false;
                    }
                }

                foreach (var field in GraphJson.FolderReadFields)
                {
                    if (!Has(GraphApp.FolderSelectFields, field))
                    {
                        return false;
                    }
                }

                // Documented: the document types may not carry a name the parser does not map.
                var declared = DocumentMemberNames();
                foreach (var name in declared)
                {
                    if (!Has(GraphJson.MessageReadFields, name)
                        && !Has(GraphJson.FolderReadFields, name)
                        && !Has(GraphJson.EnvelopeFields, name))
                    {
                        return false;
                    }
                }

                foreach (var name in GraphJson.MessageReadFields)
                {
                    if (!Has(declared, name))
                    {
                        return false;
                    }
                }

                foreach (var name in GraphJson.FolderReadFields)
                {
                    if (!Has(declared, name))
                    {
                        return false;
                    }
                }

                foreach (var name in GraphJson.EnvelopeFields)
                {
                    if (!Has(declared, name))
                    {
                        return false;
                    }
                }

                return true;
            });

            return failed;
        }

        /// <summary>
        /// The sub-members that ride along inside a field the request does ask for: "from" brings
        /// the sender's name and address, "flag" brings the flag status.
        /// </summary>
        private static readonly string[] NestedFields = { "emailAddress", "name", "address", "flagStatus" };

        /// <summary>
        /// Every json name the document types declare, read through their attributes so a renamed
        /// property cannot hide a drift between the parser and the payload.
        /// </summary>
        private static string[] DocumentMemberNames()
        {
            var names = new List<string>();
            foreach (var nested in typeof(GraphJson).GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic))
            {
                foreach (var property in nested.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
                {
                    foreach (var attribute in property.GetCustomAttributes(typeof(DataMemberAttribute), false))
                    {
                        var member = attribute as DataMemberAttribute;
                        if (member != null && !string.IsNullOrEmpty(member.Name) && !names.Contains(member.Name))
                        {
                            names.Add(member.Name);
                        }
                    }
                }
            }

            return names.ToArray();
        }

        private static bool Has(string[] values, string wanted)
        {
            return Array.IndexOf(values, wanted) >= 0;
        }

        private static bool Flagged(string status)
        {
            var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-f", status == null ? "" : ",\"flag\":{\"flagStatus\":\"" + status + "\"}")), "Inbox");
            return rows != null && rows.Count == 1 && rows[0].HasFlag;
        }

        private static int Importance(string importance)
        {
            var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-i", importance == null ? "" : ",\"importance\":\"" + importance + "\"")), "Inbox");
            return rows == null || rows.Count != 1 ? -1 : rows[0].Importance;
        }

        private static bool Meeting(string messageType)
        {
            var rows = GraphJson.ParseMessages(Page(Message("AAMkAD-m", messageType == null ? "" : ",\"meetingMessageType\":\"" + messageType + "\"")), "Inbox");
            return rows != null && rows.Count == 1 && rows[0].IsMeetingRequest;
        }

        private static string Page(string messages)
        {
            return "{\"value\":[" + messages + "]}";
        }

        private static string Message(string id, string extraFields)
        {
            return "{\"id\":\"" + id + "\"" + extraFields + "}";
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
