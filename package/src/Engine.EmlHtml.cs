// =============================================================================
//  Mailbox Transfer - engine, part 13b: text of an HTML body
// -----------------------------------------------------------------------------
//  Exchange keeps ONE body (HTML) and builds the text alternative when MIME is asked for. MapiToMime does the same when a message has an HTML
//  body and no plain text one.
// =============================================================================
using System.Text.RegularExpressions;

namespace MailboxTransfer
{
    public static class HtmlText
    {
        /// <summary>Plain text of an HTML body: tags removed, block ends become line breaks, entities decoded.</summary>
        public static string FromHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";
            string t = Regex.Replace(html, @"(?is)<(head|style|script)\b.*?</\1\s*>", "");
            t = Regex.Replace(t, @"(?i)<\s*br\s*/?\s*>", "\n");
            t = Regex.Replace(t, @"(?i)<\s*li\b[^>]*>", "\n* ");
            t = Regex.Replace(t, @"(?i)</\s*(p|div|tr|h[1-6]|ul|ol|table|blockquote)\s*>", "\n");
            t = Regex.Replace(t, @"<[^>]*>", "");
            t = System.Net.WebUtility.HtmlDecode(t).Replace(' ', ' ');
            t = Regex.Replace(t, @"[ \t]+", " ");
            t = Regex.Replace(t, @" ?\r?\n ?", "\n");
            t = Regex.Replace(t, @"\n{3,}", "\n\n");
            return t.Trim();
        }
    }
}
