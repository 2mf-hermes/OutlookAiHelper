using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public interface ISettingsStore
    {
        AppSettings Load();
        void Save(AppSettings settings);
    }
}
