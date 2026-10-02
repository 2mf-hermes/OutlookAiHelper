using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    public sealed class AiProviderProfile
    {
        public AiProviderProfile()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = "Provider";
            BaseUrl = string.Empty;
            ApiKey = string.Empty;
            Model = string.Empty;
            CachedModels = new List<string>();
        }

        public string Id { get; set; }
        public string Name { get; set; }
        public string BaseUrl { get; set; }
        public string ApiKey { get; set; }
        public string Model { get; set; }
        public List<string> CachedModels { get; set; }
    }
}
