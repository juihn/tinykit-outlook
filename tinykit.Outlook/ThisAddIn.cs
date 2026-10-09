using System;
using System.Collections.Generic;
using System.Linq;
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
        private OutlookRibbon _ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon = new OutlookRibbon(Controller);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // An exception nothing caught ends Outlook and gets the add-in disabled; at least its stack is logged.
            AppDomain.CurrentDomain.UnhandledException += (s, a) =>
            {
                try
                {
                    Log.Error("Unhandled (Outlook ends)", a.ExceptionObject as Exception);
                }
                catch
                {
                    // nothing more can be done here
                }
            };
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
                Application.ItemSend += (object item, ref bool cancel) =>
                {
                    try
                    {
                        Controller.DelaySending(item);
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Send delay", ex); // the mail still goes, without the delay
                    }
                };
            }
            catch (Exception ex)
            {
                Log.Error("ItemSend", ex);
            }

            try
            {
                Controller.SentMail = FolderToInbox.SentMail(Application, Controller.MoveSentToInbox);
                Controller.SentMail.Start();
            }
            catch (Exception ex)
            {
                Log.Error("Sent to Inbox: start", ex);
            }

            try
            {
                Controller.JunkMail = FolderToInbox.JunkMail(Application, Controller.MoveJunkToInbox);
                Controller.JunkMail.Start();
            }
            catch (Exception ex)
            {
                Log.Error("Junk to Inbox: start", ex);
            }

            try
            {
                _shortcuts = new KeyboardShortcuts();
                // The add-in's shortcuts, each on unless turned off in the Custom Shortcuts window. They work in the main
                // window only, and not while typing in a reply in its reading pane (Outlook's own keys apply there).
                const System.Windows.Forms.Keys CtrlAlt = System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt;
                AddShortcut("FindItems", CtrlAlt | System.Windows.Forms.Keys.E, "Find Items",
                    "Open (or bring to the front) the Find Items window.", ex => Controller.ShowFindItems(ex));
                AddShortcut("CustomFilter", CtrlAlt | System.Windows.Forms.Keys.F, "Custom Filter",
                    "Open the Custom Filter window for the current folder. (Outlook: Forward as attachment.)", ex => Controller.ShowCustomFilter(ex));
                AddShortcut("OpenInNewWindow", CtrlAlt | System.Windows.Forms.Keys.W, "Open in New Window",
                    "Open the current folder in a new Outlook window (Folder > Open in New Window).",
                    ex => ex.CommandBars.ExecuteMso("WebOpenInNewWindow"));
                AddShortcut("ReadingPane", CtrlAlt | System.Windows.Forms.Keys.R, "Reading Pane Right/Bottom/Off",
                    "Move the reading pane: Right, then Bottom, then Off, then Right again. (Outlook: Reply with Meeting.)",
                    CycleReadingPane);
                AddShortcut("QuickFilterBox", CtrlAlt | System.Windows.Forms.Keys.Q, "Quick Filter box",
                    "Show the TinyKit tab and put the cursor in the Quick Filter box (then Alt + a button's letter filters).",
                    FocusQuickFilterBox);
                Shortcuts.Load();
                // Tab in the ribbon skips the Quick Filter's recent-value lists.
                RibbonTabSkip.Start();
                _shortcuts.After(System.Windows.Forms.Keys.Tab, RibbonTabSkip.AfterTab);
                AddQuickFilterKeys();
            }
            catch (Exception ex)
            {
                Log.Error("Keyboard shortcuts", ex);
            }
        }

        private KeyboardShortcuts _shortcuts;

        /// <summary>The add-in's keyboard shortcuts and which are on (Custom Shortcuts window).</summary>
        internal CustomShortcuts Shortcuts { get; } = new CustomShortcuts();

        private void AddShortcut(string id, System.Windows.Forms.Keys keys, string command, string description, Action<Outlook.Explorer> run)
        {
            var shortcut = Shortcuts.Add(id, keys, command, description);
            _shortcuts.Add(keys,
                () => shortcut.Enabled && Application.ActiveWindow() is Outlook.Explorer && !KeyboardShortcuts.TypingInEditor(),
                () =>
                {
                    var explorer = Application.ActiveWindow() as Outlook.Explorer;
                    if (explorer == null)
                        return;
                    try
                    {
                        run(explorer);
                    }
                    catch (UserMessageException ex)
                    {
                        Notifier.Info(explorer, ex.Message);
                    }
                });
        }

        /// <summary>
        /// Shows the TinyKit tab and puts the focus in the Quick Filter box (looked for a few times: the tab is drawn a
        /// moment after it is shown).
        /// </summary>
        private void FocusQuickFilterBox(Outlook.Explorer explorer)
        {
            if (Controller.Focus(explorer) == null || FilterController.QuickKindsFor(Controller.Focus(explorer)).Length == 0)
                throw new UserMessageException("The Quick Filter is in mail and contact folders.");
            _ribbon.ActivateMainTab();
            var owner = WindowOwner.From(explorer);
            if (owner == null)
                return;
            RibbonFocus.FocusEditBox(owner.Handle, RibbonTabSkip.QuickBoxName, ok =>
            {
                if (!ok)
                    Log.Info("Quick Filter box: not found in the ribbon");
            });
        }

        /// <summary>
        /// In the Quick Filter box, Alt + a filter button's letter (mail: F From, S Subject, N Name, D Domain; contacts:
        /// F File As, E Email, C Company, D Department) filters by that field with the text in the box, as the button does.
        /// </summary>
        private void AddQuickFilterKeys()
        {
            var letters = ((Filtering.QuickKind[])Enum.GetValues(typeof(Filtering.QuickKind)))
                .Select(k => k.ToString()[0]).Distinct();
            foreach (var letter in letters)
            {
                var key = System.Windows.Forms.Keys.Alt | (System.Windows.Forms.Keys)char.ToUpperInvariant(letter);
                Filtering.QuickKind kind = 0;
                string text = null;
                _shortcuts.Add(key,
                    () =>
                    {
                        // checked in the hook: only in the box, and only for a letter of the folder's buttons
                        var explorer = Application.ActiveWindow() as Outlook.Explorer;
                        text = explorer == null ? null : RibbonTabSkip.QuickBoxText();
                        if (text == null)
                            return false;
                        var match = FilterController.QuickKindsFor(Controller.Focus(explorer)).Where(k => k.ToString()[0] == letter).ToList();
                        if (match.Count == 0)
                            return false;
                        kind = match[0];
                        return true;
                    },
                    () =>
                    {
                        var explorer = Application.ActiveWindow() as Outlook.Explorer;
                        if (explorer == null)
                            return;
                        try
                        {
                            Controller.SetInputText(text);
                            Controller.QuickButton(explorer, kind);
                        }
                        catch (UserMessageException ex)
                        {
                            Notifier.Info(explorer, ex.Message);
                        }
                    });
            }
        }

        // Reading pane Right → Bottom → Off → Right, through Outlook's own View > Reading Pane commands.
        private static void CycleReadingPane(Outlook.Explorer explorer)
        {
            var bars = explorer.CommandBars;
            string next = bars.GetPressedMso("ReadingPaneRight") ? "ReadingPaneBottom"
                : bars.GetPressedMso("ReadingPaneBottom") ? "ReadingPaneOff"
                : "ReadingPaneRight";
            bars.ExecuteMso(next);
        }

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
                    Controller.OnViewChanged(explorer, "startup");
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

                try
                {
                    var events = (Outlook.ExplorerEvents_10_Event)_explorer;
                    events.FolderSwitch -= OnChanged;
                    events.ViewSwitch -= OnChanged;
                    events.Close -= OnClose;
                }
                catch (Exception ex)
                {
                    Log.Info("ExplorerWatcher: events not released: " + ex.Message);
                }
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
