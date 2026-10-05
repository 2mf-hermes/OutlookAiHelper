using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Core.Security;

namespace OutlookAiHelper.Adapters.Update
{
    /// <summary>
    /// Reads the newest published release from the GitHub REST API.
    /// Only https://api.github.com/repos/{owner}/{repo}/releases/latest is ever called —
    /// no query string, no user identifier, no mail data leaves the machine here.
    /// </summary>
    public sealed class GitHubReleaseFeed : IReleaseFeed
    {
        private const int TimeoutMs = 15000;

        public ReleaseCheckResult FetchLatest(string owner, string repository)
        {
            if (string.IsNullOrEmpty(owner) || string.IsNullOrEmpty(repository))
            {
                return ReleaseCheckResult.Failure("update.error.config");
            }

            var url = "https://api.github.com/repos/" + owner + "/" + repository + "/releases/latest";
            if (!UpdateSecurityPolicy.IsTrustedApiUrl(url))
            {
                return ReleaseCheckResult.Failure("update.error.policy");
            }

            UpdateHttp.EnsureModernTls();

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.Accept = "application/vnd.github+json";
                request.UserAgent = "OutlookAiHelper/" + AppInfo.Version;
                request.Timeout = TimeoutMs;
                request.ReadWriteTimeout = TimeoutMs;
                request.AllowAutoRedirect = false;

                using (var response = (HttpWebResponse)request.GetResponse())
                using (var stream = response.GetResponseStream())
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    var release = Parse(reader.ReadToEnd());
                    if (release == null)
                    {
                        return ReleaseCheckResult.Failure("update.error.badRelease");
                    }

                    FileLogger.Info("Update check: newest release is " + release.TagName);
                    return ReleaseCheckResult.Success(release);
                }
            }
            catch (WebException ex)
            {
                var http = ex.Response as HttpWebResponse;
                if (http != null && http.StatusCode == HttpStatusCode.NotFound)
                {
                    // Repository exists but has no published release yet.
                    return ReleaseCheckResult.NoRelease();
                }

                FileLogger.Error("UpdateCheck", ex);
                return ReleaseCheckResult.Failure(
                    ex.Status == WebExceptionStatus.Timeout ? "update.error.timeout" : "update.error.network");
            }
            catch (Exception ex)
            {
                FileLogger.Error("UpdateCheck", ex);
                return ReleaseCheckResult.Failure("update.error.network");
            }
        }

        internal static ReleaseInfo Parse(string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                return null;
            }

            try
            {
                ReleaseDocument doc;
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    doc = (ReleaseDocument)new DataContractJsonSerializer(typeof(ReleaseDocument)).ReadObject(stream);
                }

                if (doc == null || string.IsNullOrEmpty(doc.TagName))
                {
                    return null;
                }

                var release = new ReleaseInfo
                {
                    TagName = doc.TagName,
                    Name = doc.Name,
                    Notes = doc.Body,
                    PageUrl = doc.HtmlUrl,
                    Draft = doc.Draft,
                    Prerelease = doc.Prerelease,
                    PublishedAt = doc.PublishedAt
                };

                if (doc.Assets != null)
                {
                    var assets = new List<ReleaseAsset>();
                    foreach (var asset in doc.Assets)
                    {
                        if (asset == null || string.IsNullOrEmpty(asset.Name))
                        {
                            continue;
                        }

                        assets.Add(new ReleaseAsset
                        {
                            Name = asset.Name,
                            DownloadUrl = asset.BrowserDownloadUrl,
                            Size = asset.Size
                        });
                    }

                    release.Assets = assets;
                }

                return release;
            }
            catch (Exception ex)
            {
                FileLogger.Error("UpdateCheck.Parse", ex);
                return null;
            }
        }

        [DataContract]
        private sealed class ReleaseDocument
        {
            [DataMember(Name = "tag_name")]
            public string TagName { get; set; }

            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "body")]
            public string Body { get; set; }

            [DataMember(Name = "html_url")]
            public string HtmlUrl { get; set; }

            [DataMember(Name = "draft")]
            public bool Draft { get; set; }

            [DataMember(Name = "prerelease")]
            public bool Prerelease { get; set; }

            [DataMember(Name = "published_at")]
            public string PublishedAt { get; set; }

            [DataMember(Name = "assets")]
            public AssetDocument[] Assets { get; set; }
        }

        [DataContract]
        private sealed class AssetDocument
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "browser_download_url")]
            public string BrowserDownloadUrl { get; set; }

            [DataMember(Name = "size")]
            public long Size { get; set; }
        }
    }
}
