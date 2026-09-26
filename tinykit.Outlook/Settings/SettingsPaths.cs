using System;
using System.IO;
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

        public static string SavedFiltersFile { get { return Path.Combine(SettingsFolder, "Saved Filters.xml"); } }

        public static string HistoryFile { get { return Path.Combine(SettingsFolder, "History.xml"); } }

        public static string ViewStateFile { get { return Path.Combine(LocalFolder, "ViewState.xml"); } }

        public static string LogFile { get { return Path.Combine(LocalFolder, "OutlookAddin.log"); } }

        /// <summary>
        /// One-time moves at startup: the pre-"Saved Filters" file name (Filters.xml), and — when the OneDrive folder has
        /// just been created — copies of the local settings into it (the local files stay as a backup).
        /// </summary>
        public static void MigrateLegacyFiles()
        {
            var legacy = Path.Combine(LocalFolder, "Filters.xml");
            var localSaved = Path.Combine(LocalFolder, "Saved Filters.xml");
            if (File.Exists(legacy) && !File.Exists(localSaved))
                File.Move(legacy, localSaved);

            if (!IsShared)
                return;
            foreach (var name in new[] { "Saved Filters.xml", "History.xml" })
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
