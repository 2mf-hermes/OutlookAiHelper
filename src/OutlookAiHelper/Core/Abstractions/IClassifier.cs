using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public interface IClassifier
    {
        ClassificationResult Classify(MailSummary mail, System.Collections.Generic.IDictionary<string, OverrideEntry> overrides);
    }
}
