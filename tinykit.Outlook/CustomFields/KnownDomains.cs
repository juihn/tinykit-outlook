using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using tinykit.OutlookAddin.Common;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Known Domains.txt: the base domains whose mail is trusted, one per line as "'yy.MM.dd(day) HH:mm:ss +zz&lt;Tab&gt;domain"
    /// (the date it was added, then the domain). A line may also be just a domain; lines starting with # are comments.
    /// A domain also covers its subdomains (fabrikam.com covers billing.fabrikam.com).
    /// The file is re-read whenever it changes, so edits made in an editor apply to the next mail.
    /// </summary>
    internal sealed class KnownDomains
    {
        private const string Header =
@"# tinykit Outlook: known domains. Mail from other domains is marked in the unknownDomain column:
#   *  the sender's domain is not known
#   +  the sender's domain is known, but a recipient's is not
#   -  the sender and all recipients are in known domains
# One domain per line, after the date it was added and a tab. A domain covers its subdomains.
# Add Known Domain (ribbon, Custom Mail Fields) adds the base domains of the selected mails' senders;
# Ctrl+click it to edit this file.
";

        private readonly string _path;
        private HashSet<string> _domains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private DateTime _loadedWrite = DateTime.MinValue;
        private long _loadedLength = -1;

        public KnownDomains(string path)
        {
            _path = path;
        }

        public string Path
        {
            get { return _path; }
        }

        /// <summary>True if the host of <paramref name="smtp"/> is a known domain or one of its subdomains.</summary>
        public bool IsKnown(string smtp)
        {
            var host = Host(smtp);
            if (host == null)
                return true; // no address to judge (e.g. system messages)
            Refresh();
            for (var h = host; h.Length > 0;)
            {
                if (_domains.Contains(h))
                    return true;
                int dot = h.IndexOf('.');
                if (dot < 0)
                    break;
                h = h.Substring(dot + 1);
            }
            return false;
        }

        /// <summary>
        /// Appends the domains not in the list yet, each with the current date. Returns the ones added.
        /// </summary>
        public List<string> Add(IEnumerable<string> domains)
        {
            Refresh();
            var added = domains.Where(d => !string.IsNullOrWhiteSpace(d))
                .Select(d => d.Trim().ToLowerInvariant())
                .Distinct()
                .Where(d => !_domains.Contains(d))
                .ToList();
            if (added.Count == 0)
                return added;

            var stamp = Stamp(DateTime.Now);
            var sb = new StringBuilder();
            if (!File.Exists(_path))
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_path));
                sb.Append(Header.Replace("\r\n", "\n").Replace("\n", "\r\n"));
            }
            else if (!EndsWithNewLine(_path))
            {
                sb.Append("\r\n");
            }
            foreach (var d in added)
                sb.Append(stamp).Append('\t').Append(d).Append("\r\n");
            File.AppendAllText(_path, sb.ToString(), new UTF8Encoding(true));
            _loadedLength = -1; // re-read on next use
            return added;
        }

        /// <summary>Creates the file with only its header comment, so it can be edited before anything is added.</summary>
        public void EnsureExists()
        {
            if (File.Exists(_path))
                return;
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_path));
            File.WriteAllText(_path, Header.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
        }

        /// <summary>
        /// "'26.09.07월 15:42:39 +08": year, month, day, Korean day of week, local time and UTC offset
        /// (hours only when whole, e.g. +08; with minutes otherwise, e.g. +05:30).
        /// </summary>
        public static string Stamp(DateTime time)
        {
            var ko = CultureInfo.GetCultureInfo("ko-KR");
            var offset = TimeZoneInfo.Local.GetUtcOffset(time);
            return "'" + time.ToString("yy.MM.dd", CultureInfo.InvariantCulture) + time.ToString("ddd", ko)
                + " " + time.ToString("HH:mm:ss", CultureInfo.InvariantCulture)
                + " " + (offset < TimeSpan.Zero ? "-" : "+") + offset.Duration().ToString(offset.Minutes == 0 ? "hh" : "hh\\:mm");
        }

        private static string Host(string smtp)
        {
            if (string.IsNullOrWhiteSpace(smtp))
                return null;
            int at = smtp.LastIndexOf('@');
            if (at < 0 || at == smtp.Length - 1)
                return null;
            return smtp.Substring(at + 1).Trim().TrimEnd('.').ToLowerInvariant();
        }

        private void Refresh()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    _domains.Clear();
                    return;
                }
                var info = new FileInfo(_path);
                if (info.LastWriteTimeUtc == _loadedWrite && info.Length == _loadedLength)
                    return;
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var raw in File.ReadAllLines(_path, Encoding.UTF8))
                {
                    var line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                        continue;
                    int tab = line.LastIndexOf('\t');
                    var domain = (tab >= 0 ? line.Substring(tab + 1) : line).Trim().TrimStart('@', '.').ToLowerInvariant();
                    if (domain.Length > 0 && domain.IndexOf(' ') < 0)
                        set.Add(domain);
                }
                _domains = set;
                _loadedWrite = info.LastWriteTimeUtc;
                _loadedLength = info.Length;
            }
            catch (IOException ex)
            {
                Log.Error("Known Domains.txt", ex); // keep the previous list (e.g. file locked while saving)
            }
        }

        private static bool EndsWithNewLine(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fs.Length == 0)
                    return true;
                fs.Seek(-1, SeekOrigin.End);
                return fs.ReadByte() == '\n';
            }
        }
    }
}
