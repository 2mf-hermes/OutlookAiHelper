namespace OutlookAiHelper.Core.Classification
{
    public sealed class ScoreReason
    {
        public ScoreReason(string code, string labelKey, int delta, string detail)
        {
            Code = code;
            LabelKey = labelKey;
            Delta = delta;
            Detail = detail;
        }

        public string Code { get; private set; }
        public string LabelKey { get; private set; }
        public int Delta { get; private set; }
        public string Detail { get; private set; }
    }
}
