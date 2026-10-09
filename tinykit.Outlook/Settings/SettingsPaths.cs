using System;
using System.IO;
using System.Linq;
using tinykit.OutlookAddin.Common;

namespace tinykit.OutlookAddin.Settings
{
    /// <summary>
    /// Where the add-in keeps its files.
    /// <list type="bullet">
    /// <item>Shared settings (Saved Filters.xml, History.xml): %OneDriveConsumer%\.config\tinykit\Outlook\ when that
    ///   folder exists, so every PC signed in to the same personal OneDrive uses the same settings; otherwise the local folder.</item>
    /// <item>Per-PC state (ViewState.xml, OutlookAddin.log): always %APPDATA%\tinykit\Outlook\.</item>
    /// </list>
    /// Word and Excel add-ins get sibling folders (…\tinykit\Word\, …\Excel\).
    /// </summary>
    internal static class SettingsPaths
    {
        private const string Product = "tinykit";
        private const string Component = "Outlook";

        /// <summary>%APPDATA%\tinykit\Outlook — per-PC files, and the settings when OneDrive sharing is not set up.</summary>
        public static string LocalFolder
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Product, Component);
            }
        }

        /// <summary>
        /// %OneDriveConsumer%\.config\tinykit\Outlook (personal OneDrive; %OneDrive% may be a work account),
        /// or null when there is no personal OneDrive.
        /// </summary>
        public static string SharedFolderCandidate
        {
            get
            {
                var oneDrive = PersonalOneDrive();
                return string.IsNullOrEmpty(oneDrive) ? null : Path.Combine(oneDrive, ".config", Product, Component);
            }
        }

        /// <summary>
        /// The personal OneDrive folder: %OneDriveConsumer%, or the OneDrive client's registry entry when the variable is
        /// missing from this process's environment.
        /// </summary>
        private static string PersonalOneDrive()
        {
            var path = Environment.GetEnvironmentVariable("OneDriveConsumer");
            if (!string.IsNullOrEmpty(path))
                return path;
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts\Personal"))
                    return key == null ? null : key.GetValue("UserFolder") as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>True when the shared (OneDrive) settings folder exists and is used.</summary>
        public static bool IsShared
        {
            get
            {
                var shared = SharedFolderCandidate;
                return shared != null && Directory.Exists(shared);
            }
        }

        /// <summary>Folder of the shared settings files (OneDrive when set up, otherwise local).</summary>
        public static string SettingsFolder
        {
            get { return IsShared ? SharedFolderCandidate : LocalFolder; }
        }

        /// <summary>Saved filters of one kind of folder: Saved Filters - Mail.xml, - Contacts.xml, - Tasks.xml.</summary>
        public static string SavedFiltersFile(ItemKind kind)
        {
            return Path.Combine(SettingsFolder, SavedFiltersName(kind));
        }

        public static string SavedFiltersName(ItemKind kind)
        {
            return "Saved Filters - " + ItemKinds.FileSuffix(kind) + ".xml";
        }

        public static string HistoryFile { get { return Path.Combine(SettingsFolder, "History.xml"); } }

        /// <summary>Base domains trusted by the domainMark column.</summary>
        public static string KnownDomainsFile { get { return Path.Combine(SettingsFolder, "Known Domains.txt"); } }

        /// <summary>Table view columns of one kind of folder: View Columns - Mail.txt, - Contacts.txt, - Tasks.txt.</summary>
        public static string ViewColumnsFile(ItemKind kind)
        {
            return Path.Combine(SettingsFolder, "View Columns - " + ItemKinds.FileSuffix(kind) + ".txt");
        }

        public static string ViewStateFile { get { return Path.Combine(LocalFolder, "ViewState.xml"); } }

        public static string LogFile { get { return Path.Combine(LocalFolder, "OutlookAddin.log"); } }

        /// <summary>The Custom Filter window's last size (this PC's screen): "width height" in pixels.</summary>
        public static string CustomFilterSizeFile { get { return Path.Combine(LocalFolder, "Custom Filter Size.txt"); } }

        /// <summary>Custom Shortcuts: which of the add-in's keyboard shortcuts are on ("id=on|off").</summary>
        public static string ShortcutsFile { get { return Path.Combine(SettingsFolder, "Custom Shortcuts.txt"); } }

        /// <summary>Find Items window: size, splitter, Message Body, item types and the accounts left out (this PC).</summary>
        public static string FindItemsSettingsFile { get { return Path.Combine(LocalFolder, "Find Items.txt"); } }

        /// <summary>
        /// One-time moves at startup: older file names (Filters.xml, then Saved Filters.xml and View Columns.txt from before
        /// the per-kind files, which were for mail), and — when the OneDrive folder has just been created — copies of
        /// the local settings into it (the local files stay as a backup).
        /// </summary>
        public static void MigrateLegacyFiles()
        {
            var mailSaved = SavedFiltersName(ItemKind.Mail);
            var mailColumns = Path.GetFileName(ViewColumnsFile(ItemKind.Mail));
            Rename(LocalFolder, "Filters.xml", "Saved Filters.xml");
            foreach (var folder in IsShared ? new[] { LocalFolder, SharedFolderCandidate } : new[] { LocalFolder })
            {
                Rename(folder, "Saved Filters.xml", mailSaved);
                Rename(folder, "View Columns.txt", mailColumns);
            }

            if (!IsShared)
                return;
            var names = ItemKinds.All.Select(SavedFiltersName)
                .Concat(ItemKinds.All.Select(k => Path.GetFileName(ViewColumnsFile(k))))
                .Concat(new[] { "History.xml", "Known Domains.txt" });
            foreach (var name in names)
            {
                var local = Path.Combine(LocalFolder, name);
                var shared = Path.Combine(SharedFolderCandidate, name);
                if (File.Exists(local) && !File.Exists(shared))
                {
                    File.Copy(local, shared);
                    Log.Info("Copied " + local + " to the shared settings folder " + shared);
                }
            }
        }

        private static void Rename(string folder, string from, string to)
        {
            var source = Path.Combine(folder, from);
            var target = Path.Combine(folder, to);
            if (File.Exists(source) && !File.Exists(target))
            {
                File.Move(source, target);
                Log.Info("Renamed " + source + " to " + to);
            }
        }

        /// <summary>Where the settings live, and how to share them, for tips and messages.</summary>
        public static string Describe()
        {
            if (IsShared)
                return "Settings: " + SettingsFolder + " (personal OneDrive, shared by every PC that uses it). "
                     + "PC-specific files stay in " + LocalFolder + ".";
            var shared = SharedFolderCandidate;
            return "Settings: " + LocalFolder + " (this PC only). "
                 + (shared == null
                     ? "To share them across PCs, sign in to a personal OneDrive and create %OneDriveConsumer%\\.config\\tinykit\\Outlook."
                     : "To share them across PCs, create " + shared + " and restart Outlook; the current settings are copied there.");
        }
    }
}
