using System;
using System.Collections.Generic;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    public sealed class ScanResultItem
    {
        public MailSummary Mail { get; set; }
        public ClassificationResult Classification { get; set; }
    }

    public sealed class ScanOutcome
    {
        public ScanOutcome()
        {
            Items = new List<ScanResultItem>();
        }

        public IList<ScanResultItem> Items { get; private set; }
        public bool Cancelled { get; set; }
        public string ErrorKey { get; set; }
    }

    public sealed class ScanMailUseCase
    {
        private readonly IMailReader _reader;
        private readonly IClassifier _classifier;
        private readonly IOverrideStore _overrides;

        public ScanMailUseCase(IMailReader reader, IClassifier classifier, IOverrideStore overrides)
        {
            _reader = reader;
            _classifier = classifier;
            _overrides = overrides;
        }

        public ScanOutcome Run(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled, Action<ScanResultItem> onItem = null)
        {
            var outcome = new ScanOutcome();
            string reason;
            if (!_reader.IsAvailable(out reason))
            {
                outcome.ErrorKey = string.IsNullOrEmpty(reason) ? "error.outlook.not_detected" : reason;
                FileLogger.Info("Scan aborted, Outlook unavailable: " + outcome.ErrorKey);
                return outcome;
            }

            var overrideMap = _overrides.Load() ?? new Dictionary<string, OverrideEntry>();
            var buffer = new List<MailSummary>();

            try
            {
                foreach (var mail in _reader.Enumerate(scope, progress, isCancelled))
                {
                    if (isCancelled != null && isCancelled())
                    {
                        outcome.Cancelled = true;
                        break;
                    }

                    buffer.Add(mail);
                }
            }
            catch (InvalidOperationException ex)
            {
                FileLogger.Error("Scan reader failed", ex);
                outcome.ErrorKey = ex.Message != null && ex.Message.StartsWith("error.", StringComparison.Ordinal)
                    ? ex.Message
                    : "error.outlook.not_detected";
                return outcome;
            }
            catch (Exception ex)
            {
                FileLogger.Error("Scan enumerate failed", ex);
                outcome.ErrorKey = "error.scan.failed";
                return outcome;
            }

            if (isCancelled != null && isCancelled() && outcome.Items.Count == 0)
            {
                outcome.Cancelled = true;
            }

            var total = buffer.Count;
            var index = 0;
            foreach (var mail in buffer)
            {
                index++;
                if (progress != null)
                {
                    progress(index, total);
                }

                try
                {
                    var item = new ScanResultItem
                    {
                        Mail = mail,
                        Classification = _classifier.Classify(mail, overrideMap)
                    };
                    outcome.Items.Add(item);
                    if (onItem != null)
                    {
                        onItem(item);
                    }
                }
                catch (Exception ex)
                {
                    FileLogger.Error("Classify failed for item", ex);
                }
            }

            return outcome;
        }
    }
}
