using OutlookAiHelper.Core.Abstractions;

namespace OutlookAiHelper.Adapters.Ai
{
    public sealed class NullAiCapability : IAiCapability
    {
        public bool IsEnabled
        {
            get { return false; }
        }

        public AiResult Suggest(AiRequest request)
        {
            return new AiResult
            {
                Success = false,
                ErrorKey = "ai.disabled"
            };
        }
    }
}
