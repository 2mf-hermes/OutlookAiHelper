using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// Maps a source kind to the reader that can serve it. Kept in one place so that
    /// adding the Graph reader is a change here and nothing in the UI moves.
    /// </summary>
    public static class MailReaderFactory
    {
        /// <summary>
        /// False for a kind this build can name but cannot read yet. The settings file is
        /// plain JSON a user can hand-edit, so a source can legitimately point at a kind
        /// with no reader; callers must fall back rather than throw.
        ///
        /// Both halves of the answer are required -- a registration the app can use and a
        /// reader that exists -- so what the UI is told can never disagree with what Create
        /// is able to hand back.
        /// </summary>
        public static bool IsImplemented(MailSourceKind kind, string graphClientId)
        {
            if (kind == MailSourceKind.LocalOutlook)
            {
                return true;
            }

            if (kind != MailSourceKind.Graph)
            {
                return false;
            }

            return GraphReaderAvailable && GraphApp.HasClientId(graphClientId);
        }

        /// <summary>
        /// The same question for callers that resolved no client id of their own, asked with
        /// the one this build shipped with.
        /// </summary>
        public static bool IsImplemented(MailSourceKind kind)
        {
            return IsImplemented(kind, GraphApp.DefaultClientId);
        }

        /// <summary>
        /// True now that the Graph reader is in this folder, and a switch rather than a scattered
        /// set of checks so that turning the cloud path off again stays a one line change.
        ///
        /// It is only half of the answer. A registration the app cannot use still means no reader,
        /// because a source the UI is told about has to be one Create can actually build -- see
        /// IsImplemented, which both the UI and Create ask.
        /// </summary>
        private static readonly bool GraphReaderAvailable = true;

        /// <summary>
        /// Returns null when the kind has no reader yet. The local reader is passed in
        /// instead of being constructed here, so the app owns exactly one Outlook COM
        /// session no matter how often a reader is resolved.
        ///
        /// A kind on its own is enough for the local mailbox and for nothing else: a cloud reader
        /// has to be told which source it serves, because the source is what says which stored
        /// sign-in belongs to it. Callers holding a source use the overload below.
        /// </summary>
        public static IMailReader Create(MailSourceKind kind, IMailReader local)
        {
            if (kind == MailSourceKind.LocalOutlook)
            {
                return local;
            }

            return null;
        }

        /// <summary>
        /// The reader for one configured source, or null when there is none this build can build:
        /// the local kind hands back the reader it was given, a cloud source is built here, and
        /// everything else -- an unknown kind, a registration the app cannot use, an id that could
        /// not name a file -- has no reader and the caller falls back to the local mailbox.
        ///
        /// The registration is a parameter rather than being read here, so that the question "can
        /// this source be read" is asked with the same value on every side: the checklist the UI is
        /// shown is computed from the same client id this method builds with, which is what stops
        /// the two from ever disagreeing.
        /// </summary>
        public static IMailReader Create(MailSourceProfile source, IMailReader local, string graphClientId)
        {
            if (source == null)
            {
                return null;
            }

            if (source.Kind == MailSourceKind.LocalOutlook)
            {
                return local;
            }

            if (!IsImplemented(source.Kind, graphClientId))
            {
                return null;
            }

            if (!JsonPaths.IsSafeSourceId(source.Id))
            {
                // Where a source's stored sign-in lives is derived from its id. The id comes out of
                // a settings file a user can hand-edit, so an id that is not a plain identifier is
                // not a source this app wrote; it is refused here rather than being turned into a
                // path or thrown out of the property above.
                return null;
            }

            return new GraphMailReader(
                graphClientId.Trim(),
                new RefreshTokenStore(JsonPaths.RefreshTokenPath(source.Id)),
                GraphTransport.Instance);
        }

        /// <summary>
        /// The same source, for callers that resolved no registration of their own: asked with the
        /// one this build shipped with.
        /// </summary>
        public static IMailReader Create(MailSourceProfile source, IMailReader local)
        {
            return Create(source, local, GraphApp.DefaultClientId);
        }
    }
}
