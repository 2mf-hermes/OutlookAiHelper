using System.Collections.Generic;

namespace OutlookAiHelper.Core.Classification
{
    public sealed class ClassificationResult
    {
        public ClassificationResult(string entryId, Quadrant quadrant, int urgencyScore, int importanceScore, IList<ScoreReason> reasons, bool fromOverride)
        {
            EntryId = entryId;
            Quadrant = quadrant;
            UrgencyScore = urgencyScore;
            ImportanceScore = importanceScore;
            Reasons = reasons;
            FromOverride = fromOverride;
        }

        public string EntryId { get; private set; }
        public Quadrant Quadrant { get; private set; }
        public int UrgencyScore { get; private set; }
        public int ImportanceScore { get; private set; }
        public IList<ScoreReason> Reasons { get; private set; }
        public bool FromOverride { get; private set; }
    }
}
