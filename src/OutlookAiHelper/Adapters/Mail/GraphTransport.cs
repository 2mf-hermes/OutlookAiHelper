using System;
using System.IO;
using System.Net;
using System.Text;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// One answer from the service, successful or not.
    ///
    /// A transport that threw would force every caller to work out whether "the service said no" or
    /// "the request never left this machine" out of an exception type. The two mean different things
    /// to the user -- sign in again, versus check the network -- so they are reported as data.
    /// </summary>
    internal sealed class GraphResponse
    {
        /// <summary>Zero when the request never completed, so there is no status to report.</summary>
        internal int StatusCode { get; set; }

        /// <summary>The response body as text, or null when there was none.</summary>
        internal string Body { get; set; }

        /// <summary>
        /// True only for a 2xx. A 401 is a refused sign-in, a 500 or an html body is a failed read:
        /// none of them is "no mail", which is why no caller here treats a failure as an empty
        /// mailbox.
        /// </summary>
        internal bool IsSuccess
        {
            get { return StatusCode >= 200 && StatusCode < 300; }
        }

        /// <summary>
        /// True when nothing came back at all: no network, no DNS, a failed TLS handshake or a
        /// timeout. Distinguished from a refusal because it is the one failure the user can fix
        /// without touching the app.
        /// </summary>
        internal bool IsOffline { get; set; }

        /// <summary>True for the two statuses that mean "the token is not usable any more".</summary>
        internal bool IsSignedOut
        {
            get { return StatusCode == 401 || StatusCode == 403; }
        }
    }

    /// <summary>
    /// The two calls this app makes to the internet on a mailbox's behalf: a GET with a bearer
    /// token, and a form POST that carries no token at all.
    ///
    /// An interface rather than a class so the reader can be checked without a network: the whole
    /// point of the offline tests is that paging, failure handling and the read-only field list are
    /// decided by code this app owns, and that code can be exercised against canned answers.
    /// </summary>
    internal interface IGraphTransport
    {
        /// <summary>
        /// A GET carrying the access token. Returns a response for every outcome -- a refused
        /// status and a dead socket are both answers, not exceptions.
        /// </summary>
        GraphResponse Get(string url, string accessToken);

        /// <summary>
        /// A form POST, used by the sign-in flow. Never carries an access token: the values in the
        /// form are the registration, the grant type and a code or refresh token.
        /// </summary>
        GraphResponse PostForm(string url, string form);
    }

    /// <summary>
    /// The real transport: HttpWebRequest, the same shape the update check already uses, because a
    /// project without a package manager has exactly one http client available to it.
    ///
    /// Three things this class is responsible for that are easy to get wrong elsewhere:
    ///  - the modern TLS version is pinned before the first request. The version a .NET process
    ///    negotiates by default is old enough that the service may refuse it outright;
    ///  - a non-2xx status arrives as an exception in this stack. Its response is the most useful
    ///    body there is -- that is where the service explains itself -- so it is read rather than
    ///    discarded;
    ///  - a body is bounded. A mailbox answer is json of a known shape, and an unbounded read of a
    ///    response from a proxy is how a mail client runs out of memory.
    /// </summary>
    internal sealed class GraphTransport : IGraphTransport
    {
        /// <summary>The one instance the app uses; the class holds no state.</summary>
        internal static readonly GraphTransport Instance = new GraphTransport();

        private const int TimeoutMilliseconds = 30000;
        private const int MaxBodyCharacters = 4 * 1024 * 1024;

        private GraphTransport()
        {
        }

        public GraphResponse Get(string url, string accessToken)
        {
            return Send(url, accessToken, null);
        }

        public GraphResponse PostForm(string url, string form)
        {
            return Send(url, null, form);
        }

        private static GraphResponse Send(string url, string accessToken, string form)
        {
            var response = new GraphResponse();

            try
            {
                // The one place in this app that opts the process into modern TLS, shared with the
                // update check rather than duplicated here.
                OutlookAiHelper.Adapters.Update.UpdateHttp.EnsureModernTls();

                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = form == null ? "GET" : "POST";
                request.Timeout = TimeoutMilliseconds;
                request.ReadWriteTimeout = TimeoutMilliseconds;
                request.AllowAutoRedirect = false;
                request.UserAgent = "OutlookAiHelper";

                if (!string.IsNullOrEmpty(accessToken))
                {
                    request.Headers[HttpRequestHeader.Authorization] = "Bearer " + accessToken;
                }

                if (form != null)
                {
                    request.ContentType = "application/x-www-form-urlencoded";
                    var bytes = Encoding.UTF8.GetBytes(form);
                    request.ContentLength = bytes.Length;
                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }

                using (var answer = (HttpWebResponse)request.GetResponse())
                {
                    response.StatusCode = (int)answer.StatusCode;
                    response.Body = ReadBody(answer);
                }

                return response;
            }
            catch (WebException ex)
            {
                var answer = ex.Response as HttpWebResponse;
                if (answer == null)
                {
                    // No response object at all: the request never got an answer from anywhere.
                    response.IsOffline = true;
                    return response;
                }

                using (answer)
                {
                    response.StatusCode = (int)answer.StatusCode;
                    response.Body = ReadBody(answer);
                }

                return response;
            }
            catch (Exception)
            {
                // A malformed url, a socket that died mid-request, a TLS refusal: all "nothing came
                // back". This runs on the UI thread of a mail client, so it is reported, not raised.
                response.IsOffline = true;
                return response;
            }
        }

        private static string ReadBody(HttpWebResponse answer)
        {
            try
            {
                using (var stream = answer.GetResponseStream())
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        var buffer = new char[8192];
                        var text = new StringBuilder();
                        int read;

                        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            text.Append(buffer, 0, read);
                            if (text.Length >= MaxBodyCharacters)
                            {
                                break;
                            }
                        }

                        return text.ToString();
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
