using System;
using OutlookAiHelper.Core.Classification;

namespace OutlookAiHelper.Core.Models
{
    public sealed class OverrideEntry
    {
        public string EntryId { get; set; }
        public Quadrant Quadrant { get; set; }
        public DateTime UpdatedOn { get; set; }
        public string Source { get; set; }

        public static OverrideEntry Manual(string entryId, Quadrant quadrant)
        {
            return new OverrideEntry
            {
                EntryId = entryId,
                Quadrant = quadrant,
                UpdatedOn = DateTime.UtcNow,
                Source = "Manual"
            };
        }
    }
}
