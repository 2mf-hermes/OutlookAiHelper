using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// The text matching both lists share. A query is split on whitespace and every word it
    /// yields has to be found somewhere in the fields the row offers: which field does not
    /// matter, and the words need not sit next to each other or in the order they were typed.
    /// That is what makes "invoice wang" find a mail from Wang about an invoice.
    ///
    /// Case is folded with the invariant culture: a list of mails is no place for a locale to
    /// change what a search means. Nothing here reads a mailbox, a store or a file - it is
    /// handed the text the app already holds, which is why searching costs no Outlook call and
    /// cannot widen what the app is allowed to see.
    /// </summary>
    public static class ListSearch
    {
        /// <summary>
        /// The words of a query, lower-cased. A query that is empty, or only spaces, yields
        /// none - and a query with no words asks for nothing, so it matches everything.
        /// </summary>
        public static string[] Tokens(string query)
        {
            if (string.IsNullOrEmpty(query))
            {
                return new string[0];
            }

            // Split on any whitespace, including the ideographic space a Chinese IME inserts.
            var parts = query.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            var words = new List<string>(parts.Length);
            foreach (var part in parts)
            {
                var word = part.Trim();
                if (word.Length > 0)
                {
                    words.Add(word.ToLowerInvariant());
                }
            }

            return words.ToArray();
        }

        /// <summary>Whether a query asks for anything at all.</summary>
        public static bool IsActive(string query)
        {
            return Tokens(query).Length > 0;
        }

        /// <summary>
        /// Whether at least one field answers the query. Every word must be found; a word may
        /// be a fragment of a longer value, so a half-typed word already narrows the list.
        /// </summary>
        public static bool Matches(string query, params string[] fields)
        {
            var words = Tokens(query);
            if (words.Length == 0)
            {
                return true;
            }

            var haystack = string.Join("\n", fields ?? new string[0]).ToLowerInvariant();
            foreach (var word in words)
            {
                if (haystack.IndexOf(word, StringComparison.Ordinal) < 0)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
