using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// Owns the answer to "which mailbox is the app reading right now?".
    ///
    /// The reader changes when the active source changes, so nothing may hold on to an
    /// IMailReader across that moment; every caller asks this object for Reader when it is
    /// about to scan. The answer is the reader for the active source: the local Outlook reader
    /// unless that source is a cloud mailbox this build can sign in to (see
    /// MailReaderFactory.Create). The indirection exists so that switching source is a settings
    /// change rather than a rewrite.
    /// </summary>
    public sealed class MailSourceRuntime
    {
        private readonly IMailReader _local;
        private readonly List<MailSourceProfile> _sources = new List<MailSourceProfile>();
        private readonly string _clientId;
        private MailSourceProfile _active;
        private IMailReader _resolved;
        private string _resolvedId;
        private string _resolvedClientId;

        public MailSourceRuntime(IMailReader local)
            : this(local, GraphApp.DefaultClientId)
        {
        }

        /// <summary>
        /// The same runtime, told which registration the app signs in with. Passed in for the same
        /// reason Create takes it: whether a cloud source can be read has to be answered with one
        /// registration everywhere, and a test has to be able to ask that on a build shipping none.
        /// </summary>
        public MailSourceRuntime(IMailReader local, string graphClientId)
        {
            _local = local;
            _clientId = graphClientId ?? string.Empty;
            _active = MailSourceProfile.Local();
        }

        /// <summary>The source every scan reads. Never null.</summary>
        public MailSourceProfile ActiveProfile
        {
            get { return _active; }
        }

        /// <summary>
        /// Every configured source, including ones this build cannot read yet, so the UI
        /// can list them and say so instead of hiding them.
        /// </summary>
        public IList<MailSourceProfile> Sources
        {
            get { return _sources; }
        }

        /// <summary>
        /// The reader for the active source. A source switch still takes effect on the next scan,
        /// because a switch goes through Configure, which forgets the reader below.
        ///
        /// The reader is remembered rather than rebuilt on every call. A cloud reader holds the
        /// access token it obtained and the one refusal it has already seen: rebuilding it on every
        /// scan would sign in again on every scan, and would forget that a sign-in was refused and
        /// so retry it forever.
        /// </summary>
        public IMailReader Reader
        {
            get
            {
                if (_resolved == null
                    || !string.Equals(_resolvedId, _active.Id, StringComparison.Ordinal)
                    || !string.Equals(_resolvedClientId, _clientId, StringComparison.Ordinal))
                {
                    _resolved = MailReaderFactory.Create(_active, _local, _clientId) ?? _local;
                    _resolvedId = _active.Id;
                    _resolvedClientId = _clientId;
                }

                return _resolved;
            }
        }

        /// <summary>
        /// Forgets the reader, so the next scan builds one again. Called when the sources or the
        /// registration change; the sign-in flow calls it too, because a reader built before a
        /// sign-in cannot know about one.
        /// </summary>
        public void InvalidateReader()
        {
            _resolved = null;
            _resolvedId = null;
            _resolvedClientId = null;
        }

        /// <summary>Whether the active source must be driven from an STA thread.</summary>
        public bool RequiresSta
        {
            get { return _active.RequiresSta; }
        }

        /// <summary>
        /// Loads the configured sources. Called after the settings load and after the user
        /// edits them. An empty list, an active id that matches nothing, a kind this build
        /// cannot read and a cloud source whose registration is missing or unusable all land
        /// on the local source, because leaving the app with no readable mailbox would look
        /// exactly like Outlook being broken.
        /// </summary>
        public void Configure(AppSettings settings)
        {
            _sources.Clear();
            InvalidateReader();

            if (settings != null && settings.MailSources != null)
            {
                foreach (var source in settings.MailSources)
                {
                    if (source != null && !string.IsNullOrEmpty(source.Id))
                    {
                        _sources.Add(source);
                    }
                }
            }

            var configured = settings == null ? null : settings.GetActiveMailSource();
            if (configured != null && MailReaderFactory.IsImplemented(configured.Kind, _clientId))
            {
                _active = configured;
                return;
            }

            var usable = FirstImplemented(_sources);
            _active = usable ?? MailSourceProfile.Local();
        }

        /// <summary>
        /// The first configured source this build can read, or null when there is none.
        ///
        /// An instance method rather than a static helper, because whether a source is readable is
        /// answered with this runtime's registration and not with the build's: a caller that was
        /// told to sign in with one client id must get the answer that goes with it.
        /// </summary>
        private MailSourceProfile FirstImplemented(IList<MailSourceProfile> sources)
        {
            foreach (var source in sources)
            {
                if (source != null && MailReaderFactory.IsImplemented(source.Kind, _clientId))
                {
                    return source;
                }
            }

            return null;
        }
    }
}
