using System;
using System.Collections.Generic;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.CustomFields;
using tinykit.OutlookAddin.Ribbon;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin
{
    public partial class ThisAddIn
    {
        internal const string Title = "tinykit – Outlook";

        private FilterController _controller;
        private Outlook.Explorers _explorers;
        private readonly List<ExplorerWatcher> _watchers = new List<ExplorerWatcher>();

        internal FilterController Controller
        {
            get { return _controller ?? (_controller = new FilterController()); }
        }

        // Called by VSTO before Startup.
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookRibbon(Controller);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _explorers = Application.Explorers;
            _explorers.NewExplorer += Watch;
            foreach (Outlook.Explorer explorer in _explorers)
                Watch(explorer);

            try
            {
                Controller.WatchSettingsFile(() => Application.ActiveExplorer());
            }
            catch (Exception ex)
            {
                Log.Error("WatchSettingsFile", ex);
            }

            try
            {
                Controller.Fields = new CustomFieldService(Application, Controller.AutoFillFields,
                    new KnownDomains(Settings.SettingsPaths.KnownDomainsFile));
                Controller.Fields.Start();
            }
            catch (Exception ex)
            {
                Log.Error("CustomFieldService.Start", ex);
            }

            try
            {
                _shortcuts = new KeyboardShortcuts();
                // Ctrl+Alt+2: Custom Filter, in the main window only (mail and item windows keep theirs, e.g. Heading 2).
                _shortcuts.Add(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.D2,
                    () => Application.ActiveWindow() is Outlook.Explorer,
                    () =>
                    {
                        var explorer = Application.ActiveWindow() as Outlook.Explorer;
                        if (explorer == null)
                            return;
                        try
                        {
                            Controller.ShowCustomFilter(explorer);
                        }
                        catch (UserMessageException ex)
                        {
                            Notifier.Info(explorer, ex.Message);
                        }
                    });
            }
            catch (Exception ex)
            {
                Log.Error("Keyboard shortcuts", ex);
            }
        }

        private KeyboardShortcuts _shortcuts;

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void Watch(Outlook.Explorer explorer)
        {
            _watchers.Add(new ExplorerWatcher(this, explorer));

            // No FolderSwitch fires for the folder an explorer opens with, so sync its view once it is shown.
            var timer = new System.Windows.Forms.Timer { Interval = 3000 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    Controller.OnViewChanged(explorer);
                }
                catch (Exception ex)
                {
                    Log.Error("Initial view sync", ex);
                }
            };
            timer.Start();
        }

        /// <summary>Keeps an explorer's event sink alive and forwards folder/view switches.</summary>
        private sealed class ExplorerWatcher
        {
            private readonly ThisAddIn _owner;
            private readonly Outlook.Explorer _explorer;

            public ExplorerWatcher(ThisAddIn owner, Outlook.Explorer explorer)
            {
                _owner = owner;
                _explorer = explorer;
                var events = (Outlook.ExplorerEvents_10_Event)explorer;
                events.FolderSwitch += OnChanged;
                events.ViewSwitch += OnChanged;
                events.Close += OnClose;
            }

            private void OnChanged()
            {
                try
                {
                    _owner.Controller.OnViewChanged(_explorer);
                }
                catch (Exception ex)
                {
                    Log.Error("ExplorerWatcher", ex);
                }
            }

            private void OnClose()
            {
                // The last window closing means Outlook is exiting (Shutdown is no longer raised): folders are still
                // reachable here, so clear the Inbox filters now if the option is on.
                try
                {
                    if (_owner.Application.Explorers.Count <= 1)
                        _owner.Controller.ClearInboxFiltersAtExit(_owner.Application, _explorer);
                }
                catch (Exception ex)
                {
                    Log.Error("Clear Inbox filters at exit", ex);
                }

                var events = (Outlook.ExplorerEvents_10_Event)_explorer;
                events.FolderSwitch -= OnChanged;
                events.ViewSwitch -= OnChanged;
                events.Close -= OnClose;
                _owner._watchers.Remove(this);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
