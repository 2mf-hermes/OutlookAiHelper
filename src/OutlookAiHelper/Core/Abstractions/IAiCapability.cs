using System.Collections.Generic;

namespace OutlookAiHelper.Core.Abstractions
{
    public sealed class AiRequest
    {
        public string Subject { get; set; }
        public string FromName { get; set; }
        public string Quadrant { get; set; }
        public IList<string> Reasons { get; set; }
        public string Instruction { get; set; }
    }

    public sealed class AiResult
    {
        public bool Success { get; set; }
        public string Text { get; set; }
        public string ErrorKey { get; set; }
    }

    public interface IAiCapability
    {
        bool IsEnabled { get; }
        AiResult Suggest(AiRequest request);
    }
}
