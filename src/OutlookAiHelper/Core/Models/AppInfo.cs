using System;
using System.Reflection;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Single source of the app's identity. The version is not hard-coded here: it is
    /// read from the assembly's informational version, which build.ps1 stamps from the
    /// VERSION file. That way the number shown in the UI can never drift from the tag
    /// of the release the updater compares against.
    /// </summary>
    public static class AppInfo
    {
        public const string DisplayName = "Outlook AI Helper";
        public const string Owner = "2mf-hermes";
        public const string Repo = "OutlookAiHelper";
        public const string BuildStamp = "1002.1";
        public const string FallbackVersion = "0.0.0";

        private static string _version;

        public static string Version
        {
            get
            {
                if (_version == null)
                {
                    _version = ReadInformationalVersion();
                }

                return _version;
            }
        }

        public static AppVersion CurrentVersion()
        {
            var parsed = AppVersion.ParseOrNull(Version);
            return parsed ?? AppVersion.ParseOrNull(FallbackVersion);
        }

        public static string DisplayVersion
        {
            get { return Version + " (" + BuildStamp + ")"; }
        }

        public static string ApiBaseUrl
        {
            get { return "https://api.github.com/repos/" + Owner + "/" + Repo; }
        }

        public static string LatestReleaseApiUrl
        {
            get { return ApiBaseUrl + "/releases/latest"; }
        }

        public static string ReleasePageUrl
        {
            get { return "https://github.com/" + Owner + "/" + Repo + "/releases"; }
        }

        public static string RepositoryUrl
        {
            get { return "https://github.com/" + Owner + "/" + Repo; }
        }

        private static string ReadInformationalVersion()
        {
            try
            {
                var assembly = typeof(AppInfo).Assembly;
                var attributes = assembly.GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false);
                if (attributes != null && attributes.Length > 0)
                {
                    var attribute = attributes[0] as AssemblyInformationalVersionAttribute;
                    if (attribute != null && !string.IsNullOrEmpty(attribute.InformationalVersion))
                    {
                        return attribute.InformationalVersion.Trim();
                    }
                }

                var name = assembly.GetName();
                if (name != null && name.Version != null)
                {
                    return name.Version.ToString(3);
                }
            }
            catch (Exception)
            {
            }

            return FallbackVersion;
        }
    }
}
