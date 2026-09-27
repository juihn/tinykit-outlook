using System;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>
    /// DASL filter strings for View.Filter / AutoFormatRule.Filter.
    /// Those properties take DASL WITHOUT the "@SQL=" prefix that Items.Restrict needs.
    /// </summary>
    internal static class Dasl
    {
        private const string PropTag = "http://schemas.microsoft.com/mapi/proptag/";
        private const string UserProperty = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        public const string Subject = "urn:schemas:httpmail:subject";
        public const string SenderSmtp = PropTag + "0x5D01001F";          // PR_SENDER_SMTP_ADDRESS
        public const string SenderEmail = PropTag + "0x0C1F001F";         // PR_SENDER_EMAIL_ADDRESS (SMTP or EX DN)
        public const string DomainRelated = UserProperty + "domainRelated";
        public const string NameRelated = UserProperty + "nameRelated";

        private const string ContactId = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/";
        public const string FileAs = "urn:schemas:contacts:fileas";
        public const string Company = "urn:schemas:contacts:o";
        public const string Department = "urn:schemas:contacts:department";
        public static readonly string[] ContactEmails =   // PidLidEmail1/2/3EmailAddress
        {
            ContactId + "8083001f", ContactId + "8093001f", ContactId + "80a3001f",
        };

        /// <summary>
        /// Builds the quick filter: the value may appear anywhere (LIKE '%value%') in
        /// From → the From (sender) e-mail address, Name → nameRelated, Subject → subject, Domain → domainRelated;
        /// FileAs, Email (any of the three addresses), Company and Department of a contact.
        /// </summary>
        public static string Build(QuickKind kind, string value)
        {
            var pattern = "%" + value + "%";
            switch (kind)
            {
                case QuickKind.From:
                    return Or(Like(SenderSmtp, pattern), Like(SenderEmail, pattern));
                case QuickKind.Name:
                    return Like(NameRelated, pattern);
                case QuickKind.Subject:
                    return Like(Subject, pattern);
                case QuickKind.Domain:
                    return Like(DomainRelated, pattern);
                case QuickKind.FileAs:
                    return Like(FileAs, pattern);
                case QuickKind.Email:
                    return Or(Array.ConvertAll(ContactEmails, p => Like(p, pattern)));
                case QuickKind.Company:
                    return Like(Company, pattern);
                case QuickKind.Department:
                    return Like(Department, pattern);
            }
            throw new ArgumentOutOfRangeException("kind");
        }

        /// <summary>
        /// True if two filters are the same text once the "@SQL=" prefix, whitespace outside literals and
        /// redundant outer parentheses are ignored (Outlook may reformat a filter it reads back).
        /// </summary>
        public static bool SameFilter(string a, string b)
        {
            return Normalize(a) == Normalize(b);
        }

        private static string Normalize(string sql)
        {
            var s = StripSqlPrefix(sql);
            var sb = new System.Text.StringBuilder(s.Length);
            bool inLiteral = false;
            foreach (var c in s)
            {
                if (c == '\'')
                    inLiteral = !inLiteral;
                if (inLiteral || !char.IsWhiteSpace(c))
                    sb.Append(inLiteral ? c : char.ToLowerInvariant(c));
            }
            s = sb.ToString();
            while (s.Length > 1 && s[0] == '(' && s[s.Length - 1] == ')' && Balanced(s.Substring(1, s.Length - 2)))
                s = s.Substring(1, s.Length - 2);
            return s;
        }

        private static bool Balanced(string s)
        {
            int depth = 0;
            bool inLiteral = false;
            foreach (var c in s)
            {
                if (c == '\'') inLiteral = !inLiteral;
                else if (!inLiteral && c == '(') depth++;
                else if (!inLiteral && c == ')' && --depth < 0) return false;
            }
            return depth == 0;
        }

        /// <summary>"property" = 'value' with the value quoted for DASL.</summary>
        public static string PropertyEquals(string property, string value)
        {
            return Eq(property, value);
        }

        /// <summary>
        /// A condition for a pattern where % stands for any text. Without % it is an equality. DASL LIKE honours %
        /// only at the start or end, so "a%b%c" becomes (LIKE 'a%' AND LIKE '%b%' AND LIKE '%c').
        /// </summary>
        public static string PropertyMatches(string property, string pattern)
        {
            if (pattern.IndexOf('%') < 0)
                return Eq(property, pattern);
            var parts = pattern.Split('%');
            var conditions = new System.Collections.Generic.List<string>();
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length == 0)
                    continue;
                bool first = i == 0, last = i == parts.Length - 1;
                conditions.Add(Like(property, (first ? "" : "%") + parts[i] + (last ? "" : "%")));
            }
            if (conditions.Count == 0)
                return Quote(property) + " IS NOT NULL"; // just "%"
            return conditions.Count == 1 ? conditions[0] : "(" + string.Join(" AND ", conditions) + ")";
        }

        private static readonly System.Text.RegularExpressions.Regex VariablePart = new System.Text.RegularExpressions.Regex(
            // numbers with their inner separators (dates, times, amounts, ids) and Korean units (9월 27일, 1,234원),
            // but not digits glued to Latin letters (names and codes such as Zero1, G12, P3B5) ...
            @"(?<![A-Za-z\d])(?>\d+(?:[.,:/\-]\d+)*)[년월일시분초원]?(?![A-Za-z\d])"
            // ... and month / weekday names (English, full or short; Korean weekdays)
            + @"|\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?"
            + @"|mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?)\b"
            + @"|[월화수목금토일]요일",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

        private static readonly System.Text.RegularExpressions.Regex AdjacentWildcards = new System.Text.RegularExpressions.Regex(
            @"%(?:[\s,./:\-()]*%)+");

        /// <summary>
        /// A subject pattern for recurring mail: numbers (dates, times, amounts, ids) and month/weekday names become %,
        /// and neighbouring ones merge: "Your trip with Gojek on Friday, 26 September" → "Your trip with Gojek on %".
        /// </summary>
        public static string SubjectPattern(string subject)
        {
            var s = (subject ?? "").Trim();
            s = VariablePart.Replace(s, "%");
            s = AdjacentWildcards.Replace(s, "%");
            return s;
        }

        /// <summary>Accepts filters written with or without the "@SQL=" prefix.</summary>
        public static string StripSqlPrefix(string sql)
        {
            if (sql == null)
                return "";
            sql = sql.Trim();
            return sql.StartsWith("@SQL=", StringComparison.OrdinalIgnoreCase) ? sql.Substring(5).Trim() : sql;
        }

        private static string Quote(string property)
        {
            return "\"" + property + "\"";
        }

        private static string Literal(string value)
        {
            return "'" + value.Replace("'", "''") + "'";
        }

        private static string Eq(string property, string value)
        {
            return Quote(property) + " = " + Literal(value);
        }

        public static string Like(string property, string pattern)
        {
            return Quote(property) + " LIKE " + Literal(pattern);
        }

        private static string Or(params string[] parts)
        {
            return parts.Length == 1 ? parts[0] : "(" + string.Join(" OR ", parts) + ")";
        }
    }
}
