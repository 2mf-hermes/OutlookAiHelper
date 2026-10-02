using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public interface IOverrideStore
    {
        IDictionary<string, OverrideEntry> Load();
        void Save(IDictionary<string, OverrideEntry> overrides);
    }
}
