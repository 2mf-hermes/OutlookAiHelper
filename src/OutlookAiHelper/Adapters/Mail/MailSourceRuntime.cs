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
    /// about to scan. In this version the answer is always the local Outlook reader,
    /// because no other kind is implemented yet (see MailReaderFactory.IsImplemented).
    /// The indirection exists so that switching source becomes a settings change rather
    /// than a rewrite.
    /// </summary>
    public sealed class MailSourceRuntime
    {
        private readonly IMailReader _local;
        private readonly List<MailSourceProfile> _sources = new List<MailSourceProfile>();
        private MailSourceProfile _active;

        public MailSourceRuntime(IMailReader local)
        {
            _local = local;
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
        /// The reader for the active source, resolved on each call so a source switch
        /// takes effect on the next scan.
        /// </summary>
        public IMailReader Reader
        {
            get
            {
                var reader = MailReaderFactory.Create(_active.Kind, _local);
                return reader ?? _local;
            }
        }

        /// <summary>Whether the active source must be driven from an STA thread.</summary>
        public bool RequiresSta
        {
            get { return _active.RequiresSta; }
        }

        /// <summary>
        /// Loads the configured sources. Called after the settings load and after the user
        /// edits them. An empty list, an active id that matches nothing and a kind this
        /// build cannot read all land on the local source, because leaving the app with no
        /// readable mailbox would look exactly like Outlook being broken.
        /// </summary>
        public void Configure(AppSettings settings)
        {
            _sources.Clear();

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
            if (configured != null && MailReaderFactory.IsImplemented(configured.Kind))
            {
                _active = configured;
                return;
            }

            var usable = FirstImplemented(_sources);
            _active = usable ?? MailSourceProfile.Local();
        }

        private static MailSourceProfile FirstImplemented(IList<MailSourceProfile> sources)
        {
            foreach (var source in sources)
            {
                if (source != null && MailReaderFactory.IsImplemented(source.Kind))
                {
                    return source;
                }
            }

            return null;
        }
    }
}
