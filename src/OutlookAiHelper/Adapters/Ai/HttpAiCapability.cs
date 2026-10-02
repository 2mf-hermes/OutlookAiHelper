using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using OutlookAiHelper.Core.Abstractions;

namespace OutlookAiHelper.Adapters.Ai
{
    public sealed class AiProviderConfig
    {
        public string BaseUrl { get; set; }
        public string Model { get; set; }
        public string ApiKey { get; set; }
        public int TimeoutMs { get; set; }
    }

    public sealed class HttpAiCapability : IAiCapability
    {
        private readonly AiProviderConfig _config;

        public HttpAiCapability(AiProviderConfig config)
        {
            _config = config ?? new AiProviderConfig();
        }

        public bool IsEnabled
        {
            get
            {
                return !string.IsNullOrEmpty(_config.BaseUrl)
                    && !string.IsNullOrEmpty(_config.Model)
                    && !string.IsNullOrEmpty(_config.ApiKey);
            }
        }

        /// <summary>OpenAI-compatible GET /v1/models</summary>
        public System.Collections.Generic.List<string> ListModels()
        {
            var models = new System.Collections.Generic.List<string>();
            if (string.IsNullOrEmpty(_config.BaseUrl) || string.IsNullOrEmpty(_config.ApiKey))
            {
                return models;
            }

            try
            {
                var url = _config.BaseUrl.TrimEnd('/') + "/models";
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.Timeout = _config.TimeoutMs > 0 ? _config.TimeoutMs : 15000;
                request.Headers["Authorization"] = "Bearer " + _config.ApiKey;
                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var json = reader.ReadToEnd();
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                    {
                        var doc = (ModelListDoc)new DataContractJsonSerializer(typeof(ModelListDoc)).ReadObject(stream);
                        if (doc != null && doc.Data != null)
                        {
                            foreach (var item in doc.Data)
                            {
                                if (item != null && !string.IsNullOrEmpty(item.Id))
                                {
                                    models.Add(item.Id);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("ListModels", ex);
            }

            return models;
        }

        public AiResult Suggest(AiRequest request)
        {
            if (!IsEnabled)
            {
                return Fail("ai.disabled");
            }

            if (request == null)
            {
                return Fail("ai.bad_request");
            }

            try
            {
                var url = _config.BaseUrl.TrimEnd('/') + "/chat/completions";
                var payload = BuildPayload(request);
                var response = Post(url, payload);
                var text = ExtractContent(response);
                if (string.IsNullOrEmpty(text))
                {
                    return Fail("ai.empty_response");
                }

                return new AiResult { Success = true, Text = text };
            }
            catch (WebException)
            {
                return Fail("ai.network_failed");
            }
            catch (TimeoutException)
            {
                return Fail("ai.timeout");
            }
            catch (Exception)
            {
                return Fail("ai.failed");
            }
        }

        private string BuildPayload(AiRequest request)
        {
            var reasons = request.Reasons == null ? string.Empty : string.Join("; ", request.Reasons);
            var user = "Subject: " + (request.Subject ?? string.Empty)
                + "\nFrom: " + (request.FromName ?? string.Empty)
                + "\nQuadrant: " + (request.Quadrant ?? string.Empty)
                + "\nReasons: " + reasons
                + "\nTask: " + (request.Instruction ?? "Briefly explain why this mail is in this quadrant and the next action.");

            // DataContract JSON keeps us dependency-free and avoids hand-rolled escaping.
            var doc = new ChatRequestDocument
            {
                Model = _config.Model,
                Messages = new[]
                {
                    new ChatMessage { Role = "system", Content = "You are a concise assistant. Do not invent mail body content." },
                    new ChatMessage { Role = "user", Content = user }
                }
            };

            using (var stream = new MemoryStream())
            {
                new DataContractJsonSerializer(typeof(ChatRequestDocument)).WriteObject(stream, doc);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private string Post(string url, string payload)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Timeout = _config.TimeoutMs > 0 ? _config.TimeoutMs : 15000;
            request.ReadWriteTimeout = request.Timeout;
            request.Headers["Authorization"] = "Bearer " + _config.ApiKey;

            using (var stream = request.GetRequestStream())
            {
                var bytes = Encoding.UTF8.GetBytes(payload);
                stream.Write(bytes, 0, bytes.Length);
            }

            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private static string ExtractContent(string responseJson)
        {
            if (string.IsNullOrEmpty(responseJson))
            {
                return null;
            }

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(responseJson)))
            {
                var doc = (ChatResponseDocument)new DataContractJsonSerializer(typeof(ChatResponseDocument)).ReadObject(stream);
                if (doc == null || doc.Choices == null || doc.Choices.Length == 0)
                {
                    return null;
                }

                var message = doc.Choices[0].Message;
                return message == null ? null : message.Content;
            }
        }

        private static AiResult Fail(string key)
        {
            return new AiResult { Success = false, ErrorKey = key };
        }

        [DataContract]
        private sealed class ModelListDoc
        {
            [DataMember(Name = "data")]
            public ModelItem[] Data { get; set; }
        }

        [DataContract]
        private sealed class ModelItem
        {
            [DataMember(Name = "id")]
            public string Id { get; set; }
        }

        [DataContract]
        private sealed class ChatRequestDocument
        {
            [DataMember(Name = "model")]
            public string Model { get; set; }

            [DataMember(Name = "messages")]
            public ChatMessage[] Messages { get; set; }
        }

        [DataContract]
        private sealed class ChatMessage
        {
            [DataMember(Name = "role")]
            public string Role { get; set; }

            [DataMember(Name = "content")]
            public string Content { get; set; }
        }

        [DataContract]
        private sealed class ChatResponseDocument
        {
            [DataMember(Name = "choices")]
            public ChatChoice[] Choices { get; set; }
        }

        [DataContract]
        private sealed class ChatChoice
        {
            [DataMember(Name = "message")]
            public ChatMessage Message { get; set; }
        }
    }
}
