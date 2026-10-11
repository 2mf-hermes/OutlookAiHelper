using System;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The cloud source's contract, checked before any of it touches the network: what this
    /// build is allowed to promise, which permission it asks for, and which fields it reads.
    ///
    /// Everything here is a statement about the shipped code, not about a live mailbox. A
    /// contract that only a signed-in account could test would not be tested at all on a build
    /// that has no registration yet.
    /// </summary>
    public static class GraphMailTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("an empty, blank or placeholder client id is refused", () =>
                !GraphApp.HasClientId(null)
                && !GraphApp.HasClientId("")
                && !GraphApp.HasClientId("   ")
                && !GraphApp.HasClientId("not-a-guid")
                && !GraphApp.HasClientId("PUT-YOUR-CLIENT-ID-HERE"));

            failed += Check("a real client id is accepted, with or without padding", () =>
                GraphApp.HasClientId("11111111-2222-3333-4444-555555555555")
                && GraphApp.HasClientId("  11111111-2222-3333-4444-555555555555  "));

            failed += Check("sign-in asks for the tenant-agnostic v2 endpoints", () =>
                GraphApp.Tenant == "common"
                && GraphApp.DeviceCodeUrl == "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
                && GraphApp.TokenUrl == "https://login.microsoftonline.com/common/oauth2/v2.0/token"
                && GraphApp.ScopeText
                    == "openid profile offline_access Mail.ReadBasic");

            failed += Check("every scope asked for is on the read-only allowlist", () =>
            {
                string[] allowed = { "openid", "profile", "offline_access", "Mail.ReadBasic" };
                if (GraphApp.Scopes.Length != allowed.Length)
                {
                    return false;
                }

                foreach (var scope in GraphApp.Scopes)
                {
                    if (Array.IndexOf(allowed, scope) < 0)
                    {
                        return false;
                    }

                    if (scope.IndexOf("Write", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return false;
                    }

                    if (scope.IndexOf("Send", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return false;
                    }

                    if (scope.IndexOf(".default", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return false;
                    }
                }

                return true;
            });

            failed += Check("the field list carries no body and no attachment", () =>
            {
                foreach (var field in GraphApp.SelectFields)
                {
                    var low = field.ToLowerInvariant();
                    if (low.IndexOf("body", StringComparison.Ordinal) >= 0)
                    {
                        return false;
                    }

                    if (low.IndexOf("attachment", StringComparison.Ordinal) >= 0)
                    {
                        return false;
                    }

                    if (low.IndexOf("unique", StringComparison.Ordinal) >= 0)
                    {
                        return false;
                    }
                }

                return true;
            });

            failed += Check("the field list covers every metadata column the mail list shows", () =>
                Has(GraphApp.SelectFields, "id")
                && Has(GraphApp.SelectFields, "subject")
                && Has(GraphApp.SelectFields, "from")
                && Has(GraphApp.SelectFields, "receivedDateTime")
                && Has(GraphApp.SelectFields, "isRead")
                && Has(GraphApp.SelectFields, "flag")
                && Has(GraphApp.SelectFields, "importance")
                && Has(GraphApp.SelectFields, "meetingMessageType"));

            failed += Check("a folder id is escaped into the message url", () =>
                GraphApp.MessagesUrl(null) == "https://graph.microsoft.com/v1.0/me/messages"
                && GraphApp.MessagesUrl("") == "https://graph.microsoft.com/v1.0/me/messages"
                && GraphApp.MessagesUrl("inbox") == "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
                && GraphApp.MessagesUrl("AAMkAD A=")
                    == "https://graph.microsoft.com/v1.0/me/mailFolders/AAMkAD%20A%3D/messages"
                && GraphApp.MailFoldersUrl == "https://graph.microsoft.com/v1.0/me/mailFolders");

            failed += Check("the UI can never be told a source is available that the factory refuses to build", () =>
                MailReaderFactory.IsImplemented(MailSourceKind.Graph, "11111111-2222-3333-4444-555555555555")
                    == (MailReaderFactory.Create(MailSourceKind.Graph, null) != null));

            failed += Check("a source with no client id is never offered, but the local mailbox always is", () =>
                !MailReaderFactory.IsImplemented(MailSourceKind.Graph, "")
                && !MailReaderFactory.IsImplemented(MailSourceKind.Graph, null)
                && MailReaderFactory.IsImplemented(MailSourceKind.LocalOutlook, "")
                && MailReaderFactory.IsImplemented(MailSourceKind.LocalOutlook, null));

            failed += Check("an unknown kind is neither offered nor built", () =>
                !MailReaderFactory.IsImplemented((MailSourceKind)99, "11111111-2222-3333-4444-555555555555")
                && MailReaderFactory.Create((MailSourceKind)99, null) == null);

            return failed;
        }

        private static bool Has(string[] values, string wanted)
        {
            return Array.IndexOf(values, wanted) >= 0;
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
