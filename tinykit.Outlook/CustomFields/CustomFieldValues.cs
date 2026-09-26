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

        public IEnumerable<KeyValuePair<string, string>> Pairs
        {
            get
            {
                yield return new KeyValuePair<string, string>(CustomFieldNames.DomainRelated, DomainRelated);
                yield return new KeyValuePair<string, string>(CustomFieldNames.NameRelated, NameRelated);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Me, Me);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Tos, Tos);
                yield return new KeyValuePair<string, string>(CustomFieldNames.Ccs, Ccs);
            }
        }

        public override string ToString()
        {
            return DomainRelated + " | " + NameRelated + " | " + Me + " | " + Tos + " | " + Ccs;
        }
    }
}
