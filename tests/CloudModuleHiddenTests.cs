using System;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The state this build ships in: the cloud module is hidden.
    ///
    /// The reader, the sign-in and the settings screen for a cloud mailbox are all in the folder, so
    /// the only thing keeping a source nobody can read off the screen is one static readonly field in
    /// MailReaderFactory. A field is easy to flip and even easier to flip back by accident, and
    /// nothing else here would notice: every other test passes either way, and the settings page
    /// would simply grow a row.
    ///
    /// What is pinned is the promise rather than the code. While the module is hidden, a registration
    /// -- even a well formed one -- makes no cloud source readable, nothing but the local mailbox is
    /// offerable, and the first-run question stays unasked; the cloud kind is still named by the build
    /// so that re-opening the module is a one line change. When the module is opened again these
    /// checks are meant to be rewritten, not deleted: what replaces them is the opposite statement,
    /// and it should be as explicit.
    /// </summary>
    public static class CloudModuleHiddenTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("a registration alone cannot make a cloud source readable while the module is hidden", () =>
                !MailReaderFactory.IsImplemented(MailSourceKind.Graph, RealLookingClientId)
                && !MailReaderFactory.IsImplemented(MailSourceKind.Graph)
                && MailReaderFactory.IsImplemented(MailSourceKind.LocalOutlook, RealLookingClientId));

            failed += Check("nothing but the local mailbox is offerable, so the add control has nothing to open", () =>
                !MailSourceAvailability.CanOfferNew(MailReaderFactory.IsImplemented)
                && MailSourceAvailability.FirstOfferableKind(MailReaderFactory.IsImplemented, MailSourceAvailability.CloudKinds) == null);

            failed += Check("a build with nothing to ask about never opens the first-run question", () =>
                !MailSourceAvailability.NeedsFirstRunChoice(MailReaderFactory.IsImplemented, false));

            failed += Check("the source list is the local mailbox on its own", () =>
            {
                var shown = MailSourceAvailability.DisplayList(null);

                return shown != null
                    && shown.Count == 1
                    && shown[0].Kind == MailSourceKind.LocalOutlook
                    && shown[0].Id == MailSourceProfile.LocalId;
            });

            failed += Check("the cloud kind is still named by this build, so re-opening the module is one line", () =>
                Array.IndexOf(MailSourceAvailability.CloudKinds, MailSourceKind.Graph) >= 0);

            return failed;
        }

        /// <summary>
        /// A registration of the shape GraphApp accepts, so that the checks above fail for the reason
        /// they name -- the switch -- rather than because the client id looked like a placeholder.
        /// </summary>
        private const string RealLookingClientId = "11111111-2222-3333-4444-555555555555";

        private static int Check(string name, Func<bool> body)
        {
            bool ok;
            try
            {
                ok = body();
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name);
                Console.WriteLine("   " + ex.GetType().Name + ": " + ex.Message);
                return 1;
            }

            Console.WriteLine((ok ? "PASS " : "FAIL ") + name);
            return ok ? 0 : 1;
        }
    }
}
