using System.Collections.Generic;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>User-property (custom column) names written on mail items.</summary>
    internal static class CustomFieldNames
    {
        public const string DomainRelated = "domainRelated";
        public const string NameRelated = "nameRelated";
        public const string Me = "me";
        public const string Tos = "tos";
        public const string Ccs = "ccs";
        public const string DomainMark = "domainMark";

        /// <summary>domainMark's name before 2026-10 (values "*", "+", "-"); moved to domainMark at startup.</summary>
        public const string LegacyUnknownDomain = "unknownDomain";

        public static readonly string[] All = { DomainRelated, NameRelated, Me, Tos, Ccs, DomainMark };

        /// <summary>DASL name of a (PS_PUBLIC_STRINGS) user property, for filters.</summary>
        public static string Dasl(string name)
        {
            return "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + name;
        }
    }

    /// <summary>Computed custom column values of one item. Every value is text; "-" means none.</summary>
    internal sealed class CustomFieldValues
    {
        public const string None = "-";

        public string DomainRelated = None;
        public string NameRelated = None;
        public string Me = None;
        public string Tos = None;
        public string Ccs = None;

        /// <summary>
        /// "🅄" unknown sender domain, "+" known sender but an unknown recipient domain, "-" all known; "🅙" moved from
        /// Junk Email to the Inbox (Move Junk Mail to Inbox), which recomputing keeps.
        /// </summary>
        public string DomainMark = None;

        public const string UnknownSender = "🅄";    // 🅄 SQUARED LATIN CAPITAL LETTER U
        public const string UnknownRecipient = "+";
        public const string FromJunk = "🅙";         // 🅙 NEGATIVE CIRCLED LATIN CAPITAL LETTER J
        public const string LegacyUnknownSender = "*";       // unknownDomain's mark before 2026-10

        public IEnumerable<KeyValuePair<string, string>> Pairs
        {
            get
            {
                yield return new KeyValuePair<string, string>(CustomFieldNames.DomainRelated, DomainRelated);
                yield return new KeyValuePair<string, string>(CustomFieldNames.NameRelated, NameRelated);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Me, Me);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Tos, Tos);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Ccs, Ccs);
                yield return new KeyValuePair<string, string>(CustomFieldNames.DomainMark, DomainMark);
            }
        }

        public override string ToString()
        {
            return DomainRelated + " | " + NameRelated + " | " + Me + " | " + Tos + " | " + Ccs + " | " + DomainMark;
        }
    }
}
