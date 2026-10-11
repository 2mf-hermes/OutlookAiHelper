using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// The one place a response body becomes a [DataContract] object.
    ///
    /// The mailbox payloads and the sign-in answers both come through here, because the rule they
    /// share is the rule that matters: a body that is not a json object is a failed read rather
    /// than an empty answer. The case that rule exists for is an html sign-in page handed back by
    /// a corporate proxy -- reporting "no mail" for what is really "the request never reached the
    /// mailbox" is the one outcome this app must not have.
    ///
    /// Two copies of a rule like this are two rules, so it lives here and both callers use it.
    /// </summary>
    internal static class JsonBody
    {
        /// <summary>
        /// Reads one object out of the text, or null when the text is not one. Null always means
        /// "this read failed"; an object that parses but carries nothing useful is a different
        /// thing, and describing it is left to the caller.
        /// </summary>
        internal static T Read<T>(string json) where T : class
        {
            if (string.IsNullOrEmpty(json))
            {
                return null;
            }

            var trimmed = json.TrimStart();
            if (trimmed.Length == 0 || trimmed[0] != '{')
            {
                // Every body this app reads is a json object. Anything else -- an html sign-in
                // page handed back by a proxy, a bare array, a quoted string -- is a failed read,
                // and saying so is better than reporting an empty mailbox.
                return null;
            }

            try
            {
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
                }
            }
            catch (Exception)
            {
                // A malformed or truncated payload is a failed read, not a crash: this runs on the
                // UI thread of a mail client.
                return null;
            }
        }
    }
}
