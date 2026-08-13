using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;
using Microsoft.Win32;
using UglyToad.PdfPig;

namespace Datenfinder.UI
{
    public partial class MainWindow : Window
    {
        private const int MailItemClass = 43;
        private const int MaximumSearchResults = 500;

        /*
         * Build 1120
         *
         * Indexschema bleibt 1110.
         *
         * Neue Funktionen:
         * - Mehrfachauswahl
         * - MSG-Export
         * - Mailketten/Sachverhalte über ConversationID
         * - Gesamte Mailkette exportieren
         */
        private const string IndexSchema = "1110";
        private const string PreviousIndexSchema = "1060";

        private const long MaximumAttachmentBytes =
            50L * 1024L * 1024L;

        private const int MaximumCharactersPerAttachment =
            250000;

        private const int MaximumAttachmentCharactersPerMail =
            750000;

        private readonly string _indexFolder;
        private readonly string _indexPath;
        private readonly string _syncInfoPath;

        private readonly DispatcherTimer _automaticUpdateTimer;

        private bool _updateInProgress;

        private int _totalFolderCount;
        private int _totalMailCount;
        private int _processedMailCount;

        private int _incrementalNewCount;
        private int _incrementalChangedCount;

        private int _incrementalFolderTotal;
        private int _incrementalFolderProcessed;

        public MainWindow()
        {
            InitializeComponent();

            _indexFolder =
                Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.MyDocuments),
                    "Datenfinder Gemeinde Reichenau PRO");

            _indexPath =
                Path.Combine(
                    _indexFolder,
                    "Outlook-Inhaltsindex.txt");

            _syncInfoPath =
                Path.Combine(
                    _indexFolder,
                    "Outlook-Sync.txt");

            InitializeFilters();
            CheckExistingIndex();

            _automaticUpdateTimer =
                new DispatcherTimer
                {
                    Interval =
                        TimeSpan.FromHours(1)
                };

            _automaticUpdateTimer.Tick +=
                AutomaticUpdateTimer_Tick;

            _automaticUpdateTimer.Start();
        }

        // =========================================================
        // START / FILTER
        // =========================================================

        private void InitializeFilters()
        {
            AttachmentComboBox.SelectedIndex = 0;
            FlagComboBox.SelectedIndex = 0;
            SortComboBox.SelectedIndex = 0;

            MailboxComboBox.Items.Clear();

            MailboxComboBox.Items.Add(
                "Alle Postfächer");

            MailboxComboBox.SelectedIndex = 0;

            SenderFilterTextBox.Text = "";
            RecipientFilterTextBox.Text = "";

            SearchSubjectCheckBox.IsChecked = true;
            SearchBodyCheckBox.IsChecked = true;
            SearchAttachmentsCheckBox.IsChecked = false;

            ActiveFiltersText.Text =
                "Aktive Filter: keine";

            SelectionCountText.Text =
                "Keine Mail ausgewählt";
        }

        private void CheckExistingIndex()
        {
            if (!File.Exists(_indexPath))
            {
                SearchButton.IsEnabled = false;

                CreateIndexButton.Content =
                    "Index erstellen";

                IndexStatusText.Text =
                    "Noch kein Suchindex vorhanden";

                IndexDetailsText.Text =
                    "Outlook prüfen und anschließend den ersten Index erstellen.";

                LastUpdateText.Text = "";

                return;
            }

            string schema =
                GetIndexSchema();

            if (!IsSearchableIndex())
            {
                SearchButton.IsEnabled = false;

                IndexStatusText.Text =
                    "Vorhandener Suchindex ist nicht kompatibel";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            180,
                            100,
                            0));

                IndexDetailsText.Text =
                    "Der Index muss vollständig neu erstellt werden.";

                CreateIndexButton.Content =
                    "Index erstellen";

                return;
            }

            SearchButton.IsEnabled = true;

            FileInfo fileInfo =
                new FileInfo(
                    _indexPath);

            if (schema ==
                PreviousIndexSchema)
            {
                IndexStatusText.Text =
                    "Vorhandener Index kann durchsucht werden";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            180,
                            100,
                            0));

                IndexDetailsText.Text =
                    "Für die Anhangsuche muss der Index einmal neu aufgebaut werden.";

                CreateIndexButton.Content =
                    "Anhangindex neu aufbauen";
            }
            else
            {
                IndexStatusText.Text =
                    "Outlook-Inhaltsindex ist bereit";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            0,
                            120,
                            70));

                IndexDetailsText.Text =
                    $"Indexgröße: {FormatFileSize(fileInfo.Length)}";

                CreateIndexButton.Content =
                    "Jetzt aktualisieren";
            }

            SearchStatusText.Text =
                "Suchbegriff eingeben oder Filter verwenden.";

            UpdateLastSyncDisplay();
            LoadMailboxesFromIndex();
        }

        private string GetIndexSchema()
        {
            try
            {
                using StreamReader reader =
                    new StreamReader(
                        _indexPath,
                        Encoding.UTF8,
                        true);

                for (int i = 0;
                     i < 15;
                     i++)
                {
                    string? line =
                        reader.ReadLine();

                    if (line == null)
                    {
                        break;
                    }

                    if (line.StartsWith(
                        "Schema:",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return line
                            .Substring(
                                "Schema:".Length)
                            .Trim();
                    }
                }
            }
            catch
            {
            }

            return "";
        }

        private bool IsSearchableIndex()
        {
            if (!File.Exists(_indexPath))
            {
                return false;
            }

            string schema =
                GetIndexSchema();

            return
                schema == IndexSchema ||
                schema == PreviousIndexSchema;
        }

        private bool HasAttachmentIndex()
        {
            return
                File.Exists(_indexPath) &&
                GetIndexSchema() ==
                IndexSchema;
        }

        private void LoadMailboxesFromIndex()
        {
            try
            {
                HashSet<string> mailboxes =
                    new HashSet<string>(
                        StringComparer.OrdinalIgnoreCase);

                using StreamReader reader =
                    new StreamReader(
                        _indexPath,
                        Encoding.UTF8,
                        true);

                string? line;

                while ((line =
                    reader.ReadLine()) != null)
                {
                    string[] columns =
                        line.Split('\t');

                    if (columns.Length < 14)
                    {
                        continue;
                    }

                    if (!int.TryParse(
                        columns[0],
                        out _))
                    {
                        continue;
                    }

                    string mailbox =
                        columns[5].Trim();

                    if (!string.IsNullOrWhiteSpace(
                        mailbox))
                    {
                        mailboxes.Add(
                            mailbox);
                    }
                }

                string? selected =
                    MailboxComboBox
                        .SelectedItem
                        ?.ToString();

                MailboxComboBox.Items.Clear();

                MailboxComboBox.Items.Add(
                    "Alle Postfächer");

                foreach (
                    string mailbox
                    in mailboxes.OrderBy(
                        x => x))
                {
                    MailboxComboBox.Items.Add(
                        mailbox);
                }

                if (!string.IsNullOrWhiteSpace(
                        selected) &&
                    MailboxComboBox.Items.Contains(
                        selected))
                {
                    MailboxComboBox.SelectedItem =
                        selected;
                }
                else
                {
                    MailboxComboBox.SelectedIndex =
                        0;
                }
            }
            catch
            {
                MailboxComboBox.Items.Clear();

                MailboxComboBox.Items.Add(
                    "Alle Postfächer");

                MailboxComboBox.SelectedIndex =
                    0;
            }
        }

        // =========================================================
        // SYNC
        // =========================================================

        private void UpdateLastSyncDisplay()
        {
            DateTime? lastSync =
                ReadLastSync();

            if (lastSync.HasValue)
            {
                LastUpdateText.Text =
                    $"Zuletzt aktualisiert: {lastSync.Value:dd.MM.yyyy HH:mm}";
            }
            else
            {
                LastUpdateText.Text =
                    "Noch keine automatische Aktualisierung.";
            }
        }

        private DateTime? ReadLastSync()
        {
            try
            {
                if (!File.Exists(
                    _syncInfoPath))
                {
                    return null;
                }

                string text =
                    File.ReadAllText(
                        _syncInfoPath,
                        Encoding.UTF8)
                    .Trim();

                if (DateTime.TryParseExact(
                    text,
                    "O",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out DateTime result))
                {
                    return result;
                }
            }
            catch
            {
            }

            return null;
        }

        private void WriteLastSync(
            DateTime dateTime)
        {
            Directory.CreateDirectory(
                _indexFolder);

            File.WriteAllText(
                _syncInfoPath,
                dateTime.ToString(
                    "O",
                    CultureInfo.InvariantCulture),
                Encoding.UTF8);
        }

        // =========================================================
        // AUTOMATISCHE AKTUALISIERUNG
        // =========================================================

        private async void AutomaticUpdateTimer_Tick(
            object? sender,
            EventArgs e)
        {
            if (_updateInProgress)
            {
                return;
            }

            if (!HasAttachmentIndex())
            {
                return;
            }

            await UpdateIndexIncrementallyAsync(
                false);
        }

        // =========================================================
        // OUTLOOK PRÜFEN
        // =========================================================

        private void ConnectOutlookButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            try
            {
                OutlookStatusText.Text =
                    "Status: Outlook wird geprüft ...";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            85,
                            85,
                            85));

                OutlookDetailsText.Text = "";
                MailboxNamesText.Text = "";

                CreateIndexButton.IsEnabled =
                    false;

                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType == null)
                {
                    throw new InvalidOperationException(
                        "Das klassische Microsoft Outlook wurde auf diesem PC nicht gefunden.");
                }

                outlookApplication =
                    Activator.CreateInstance(
                        outlookType);

                if (outlookApplication == null)
                {
                    throw new InvalidOperationException(
                        "Outlook konnte nicht gestartet werden.");
                }

                dynamic outlook =
                    outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace(
                        "MAPI");

                if (outlookNamespace == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-MAPI-Schnittstelle konnte nicht geöffnet werden.");
                }

                dynamic outlookNs =
                    outlookNamespace;

                stores =
                    outlookNs.Stores;

                if (stores == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-Datenspeicher konnten nicht gelesen werden.");
                }

                dynamic outlookStores =
                    stores;

                int storeCount =
                    outlookStores.Count;

                List<string> storeNames =
                    new List<string>();

                for (int i = 1;
                     i <= storeCount;
                     i++)
                {
                    object? storeObject =
                        null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(
                                i);

                        dynamic store =
                            storeObject;

                        string name =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (!string.IsNullOrWhiteSpace(
                            name))
                        {
                            storeNames.Add(
                                name);
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        ReleaseComObject(
                            storeObject);
                    }
                }

                OutlookStatusText.Text =
                    "Status: Outlook erfolgreich verbunden";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            0,
                            120,
                            70));

                OutlookDetailsText.Text =
                    $"Gefundene Postfächer: {storeCount}";

                MailboxNamesText.Text =
                    storeNames.Count > 0
                        ? "Verbunden: " +
                          string.Join(
                              "  •  ",
                              storeNames)
                        : "";

                CreateIndexButton.IsEnabled =
                    true;

                string schema =
                    GetIndexSchema();

                if (!IsSearchableIndex())
                {
                    IndexStatusText.Text =
                        "Erster Suchindex muss erstellt werden";

                    IndexDetailsText.Text =
                        "Beim ersten Lauf werden E-Mails und unterstützte Anhänge indiziert.";

                    CreateIndexButton.Content =
                        "Index erstellen";

                    SearchButton.IsEnabled =
                        false;
                }
                else if (schema ==
                    PreviousIndexSchema)
                {
                    IndexStatusText.Text =
                        "Bestehender Mailindex erkannt";

                    IndexStatusText.Foreground =
                        new SolidColorBrush(
                            Color.FromRgb(
                                180,
                                100,
                                0));

                    IndexDetailsText.Text =
                        "Betreff und Mailtext sind durchsuchbar. Für Anhänge ist einmalig ein Neuaufbau erforderlich.";

                    CreateIndexButton.Content =
                        "Anhangindex neu aufbauen";

                    SearchButton.IsEnabled =
                        true;
                }
                else
                {
                    IndexStatusText.Text =
                        "Index vorhanden – Aktualisierung bereit";

                    IndexStatusText.Foreground =
                        new SolidColorBrush(
                            Color.FromRgb(
                                0,
                                120,
                                70));

                    IndexDetailsText.Text =
                        "Nur neue oder geänderte Nachrichten werden neu eingelesen.";

                    CreateIndexButton.Content =
                        "Jetzt aktualisieren";

                    SearchButton.IsEnabled =
                        true;
                }
            }
            catch (Exception ex)
            {
                OutlookStatusText.Text =
                    "Status: Outlook-Verbindung fehlgeschlagen";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            180,
                            40,
                            40));

                OutlookDetailsText.Text =
                    ex.Message;

                CreateIndexButton.IsEnabled =
                    false;
            }
            finally
            {
                ReleaseComObject(
                    stores);

                ReleaseComObject(
                    outlookNamespace);

                ReleaseComObject(
                    outlookApplication);
            }
        }

        // =========================================================
        // INDEX BUTTON
        // =========================================================

        private async void CreateIndexButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (_updateInProgress)
            {
                return;
            }

            if (HasAttachmentIndex())
            {
                await UpdateIndexIncrementallyAsync(
                    true);
            }
            else
            {
                await CreateFullIndexAsync();
            }
        }

        // =========================================================
        // INKREMENTELLE AKTUALISIERUNG
        // =========================================================

        private async Task UpdateIndexIncrementallyAsync(
            bool userRequested)
        {
            if (_updateInProgress)
            {
                return;
            }

            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            _updateInProgress = true;

            try
            {
                CreateIndexButton.IsEnabled =
                    false;

                ConnectOutlookButton.IsEnabled =
                    false;

                ProgressPanel.Visibility =
                    Visibility.Visible;

                SetProgress(
                    0,
                    "Aktualisierung wird vorbereitet");

                ProgressCountText.Text =
                    "Vorhandener Index wird geladen ...";

                ProgressFolderText.Text = "";

                Dictionary<string, IndexRecord> records =
                    LoadExistingIndexRecords();

                int oldCount =
                    records.Count;

                DateTime lastSync =
                    ReadLastSync()
                    ?? GetBestInitialSyncPoint();

                DateTime scanSince =
                    lastSync.AddMinutes(
                        -10);

                _incrementalNewCount = 0;
                _incrementalChangedCount = 0;
                _incrementalFolderTotal = 0;
                _incrementalFolderProcessed = 0;

                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType == null)
                {
                    throw new InvalidOperationException(
                        "Das klassische Microsoft Outlook wurde nicht gefunden.");
                }

                outlookApplication =
                    Activator.CreateInstance(
                        outlookType);

                if (outlookApplication == null)
                {
                    throw new InvalidOperationException(
                        "Outlook konnte nicht gestartet werden.");
                }

                dynamic outlook =
                    outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace(
                        "MAPI");

                dynamic outlookNs =
                    outlookNamespace;

                stores =
                    outlookNs.Stores;

                dynamic outlookStores =
                    stores;

                int storeCount =
                    outlookStores.Count;

                List<string> storeNames =
                    new List<string>();

                // Phase 1: Ordner zählen

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject =
                        null;

                    object? rootFolderObject =
                        null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(
                                storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(
                            storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

                        if (!storeNames.Contains(
                            storeName))
                        {
                            storeNames.Add(
                                storeName);
                        }

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject != null)
                        {
                            _incrementalFolderTotal +=
                                CountFoldersOnly(
                                    rootFolderObject);
                        }

                        SetProgress(
                            ((double)storeIndex /
                             Math.Max(
                                 1,
                                 storeCount)) *
                            10.0,
                            $"Phase 1 von 3 – Postfach {storeIndex} von {storeCount} wird vorbereitet");

                        await RefreshUi();
                    }
                    finally
                    {
                        ReleaseComObject(
                            rootFolderObject);

                        ReleaseComObject(
                            storeObject);
                    }
                }

                if (_incrementalFolderTotal <=
                    0)
                {
                    _incrementalFolderTotal =
                        1;
                }

                // Phase 2

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject =
                        null;

                    object? rootFolderObject =
                        null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(
                                storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(
                            storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject !=
                            null)
                        {
                            await UpdateFolderIncrementallyAsync(
                                rootFolderObject,
                                storeName,
                                "",
                                scanSince,
                                records);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(
                            rootFolderObject);

                        ReleaseComObject(
                            storeObject);
                    }
                }

                // Phase 3

                SetProgress(
                    95,
                    "Phase 3 von 3 – Index wird gespeichert");

                ProgressCountText.Text =
                    $"{_incrementalNewCount:N0} neu | " +
                    $"{_incrementalChangedCount:N0} geändert";

                await RefreshUi();

                DateTime syncTime =
                    DateTime.Now;

                await WriteIndexRecordsAsync(
                    records.Values,
                    storeCount,
                    syncTime);

                WriteLastSync(
                    syncTime);

                MailboxNamesText.Text =
                    storeNames.Count > 0
                        ? "Verbunden: " +
                          string.Join(
                              "  •  ",
                              storeNames)
                        : "";

                SetProgress(
                    100,
                    "Fertig – Outlook-Inhaltsindex ist aktuell");

                ProgressCountText.Text =
                    $"{_incrementalNewCount:N0} neue | " +
                    $"{_incrementalChangedCount:N0} geänderte E-Mails";

                ProgressFolderText.Text =
                    $"Index enthält jetzt {records.Count:N0} E-Mails.";

                IndexStatusText.Text =
                    "Outlook-Inhaltsindex einschließlich Anhänge ist aktuell";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            0,
                            120,
                            70));

                IndexDetailsText.Text =
                    $"{storeCount} Postfächer | " +
                    $"{records.Count:N0} E-Mails | vorher {oldCount:N0}";

                SearchButton.IsEnabled =
                    true;

                CreateIndexButton.Content =
                    "Jetzt aktualisieren";

                LoadMailboxesFromIndex();
                UpdateLastSyncDisplay();

                if (userRequested)
                {
                    SearchStatusText.Text =
                        _incrementalNewCount == 0 &&
                        _incrementalChangedCount == 0
                            ? "Index ist bereits aktuell."
                            : $"Index aktualisiert: {_incrementalNewCount:N0} neu, " +
                              $"{_incrementalChangedCount:N0} geändert.";
                }

                await Task.Delay(
                    1200);

                ProgressPanel.Visibility =
                    Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                SetFailedProgress(
                    "Aktualisierung abgebrochen",
                    ex.Message);

                IndexStatusText.Text =
                    "Index-Aktualisierung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            180,
                            40,
                            40));

                IndexDetailsText.Text =
                    ex.Message;

                SearchButton.IsEnabled =
                    IsSearchableIndex();
            }
            finally
            {
                ReleaseComObject(
                    stores);

                ReleaseComObject(
                    outlookNamespace);

                ReleaseComObject(
                    outlookApplication);

                CreateIndexButton.IsEnabled =
                    true;

                ConnectOutlookButton.IsEnabled =
                    true;

                _updateInProgress =
                    false;
            }
        }

        private int CountFoldersOnly(
            object folderObject)
        {
            object? foldersObject =
                null;

            int result =
                1;

            try
            {
                dynamic folder =
                    folderObject;

                foldersObject =
                    folder.Folders;

                if (foldersObject ==
                    null)
                {
                    return result;
                }

                dynamic folders =
                    foldersObject;

                int count =
                    folders.Count;

                for (int i = 1;
                     i <= count;
                     i++)
                {
                    object? subFolderObject =
                        null;

                    try
                    {
                        subFolderObject =
                            folders.Item(i);

                        if (subFolderObject !=
                            null)
                        {
                            result +=
                                CountFoldersOnly(
                                    subFolderObject);
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        ReleaseComObject(
                            subFolderObject);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(
                    foldersObject);
            }

            return result;
        }

        private DateTime GetBestInitialSyncPoint()
        {
            try
            {
                FileInfo info =
                    new FileInfo(
                        _indexPath);

                return info.LastWriteTime;
            }
            catch
            {
                return
                    DateTime.Now.AddDays(
                        -1);
            }
        }

        private async Task UpdateFolderIncrementallyAsync(
            object folderObject,
            string storeName,
            string parentPath,
            DateTime scanSince,
            Dictionary<string, IndexRecord> records)
        {
            object? itemsObject =
                null;

            object? foldersObject =
                null;

            try
            {
                dynamic folder =
                    folderObject;

                string folderName =
                    SafeDynamicString(
                        folder,
                        "Name");

                if (string.IsNullOrWhiteSpace(
                    folderName))
                {
                    folderName =
                        "Unbekannter Ordner";
                }

                string currentPath =
                    string.IsNullOrWhiteSpace(
                        parentPath)
                        ? folderName
                        : parentPath +
                          " > " +
                          folderName;

                _incrementalFolderProcessed++;

                double folderProgress =
                    10.0 +
                    ((double)_incrementalFolderProcessed /
                     Math.Max(
                         1,
                         _incrementalFolderTotal)) *
                    80.0;

                SetProgress(
                    folderProgress,
                    $"Phase 2 von 3 – Ordner {_incrementalFolderProcessed:N0} von {_incrementalFolderTotal:N0}");

                ProgressFolderText.Text =
                    $"{storeName} > {currentPath}";

                ProgressCountText.Text =
                    $"{_incrementalNewCount:N0} neu | " +
                    $"{_incrementalChangedCount:N0} geändert";

                await RefreshUi();

                try
                {
                    itemsObject =
                        folder.Items;

                    if (itemsObject !=
                        null)
                    {
                        dynamic items =
                            itemsObject;

                        try
                        {
                            items.Sort(
                                "[LastModificationTime]",
                                true);
                        }
                        catch
                        {
                        }

                        int itemCount =
                            items.Count;

                        for (int i = 1;
                             i <= itemCount;
                             i++)
                        {
                            object? itemObject =
                                null;

                            try
                            {
                                itemObject =
                                    items.Item(i);

                                if (itemObject ==
                                    null)
                                {
                                    continue;
                                }

                                dynamic item =
                                    itemObject;

                                int itemClass;

                                try
                                {
                                    itemClass =
                                        item.Class;
                                }
                                catch
                                {
                                    continue;
                                }

                                if (itemClass !=
                                    MailItemClass)
                                {
                                    continue;
                                }

                                DateTime? modified =
                                    SafeLastModificationTime(
                                        item);

                                if (modified.HasValue &&
                                    modified.Value <
                                    scanSince)
                                {
                                    break;
                                }

                                IndexRecord record =
                                    BuildIndexRecord(
                                        item,
                                        storeName,
                                        currentPath);

                                if (string.IsNullOrWhiteSpace(
                                    record.EntryId))
                                {
                                    continue;
                                }

                                if (records.TryGetValue(
                                    record.EntryId,
                                    out IndexRecord?
                                        existing))
                                {
                                    if (!existing.ContentEquals(
                                        record))
                                    {
                                        records[
                                            record.EntryId] =
                                            record;

                                        _incrementalChangedCount++;
                                    }
                                }
                                else
                                {
                                    records[
                                        record.EntryId] =
                                        record;

                                    _incrementalNewCount++;
                                }
                            }
                            catch
                            {
                            }
                            finally
                            {
                                ReleaseComObject(
                                    itemObject);
                            }
                        }
                    }
                }
                catch
                {
                }

                foldersObject =
                    folder.Folders;

                if (foldersObject !=
                    null)
                {
                    dynamic folders =
                        foldersObject;

                    int subFolderCount =
                        folders.Count;

                    for (int i = 1;
                         i <= subFolderCount;
                         i++)
                    {
                        object? subFolderObject =
                            null;

                        try
                        {
                            subFolderObject =
                                folders.Item(i);

                            if (subFolderObject !=
                                null)
                            {
                                await UpdateFolderIncrementallyAsync(
                                    subFolderObject,
                                    storeName,
                                    currentPath,
                                    scanSince,
                                    records);
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            ReleaseComObject(
                                subFolderObject);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(
                    itemsObject);

                ReleaseComObject(
                    foldersObject);
            }
        }

        // =========================================================
        // VOLLINDEX
        // =========================================================

        private async Task CreateFullIndexAsync()
        {
            object? outlookApplication =
                null;

            object? outlookNamespace =
                null;

            object? stores =
                null;

            _updateInProgress =
                true;

            try
            {
                CreateIndexButton.IsEnabled =
                    false;

                ConnectOutlookButton.IsEnabled =
                    false;

                SearchButton.IsEnabled =
                    false;

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                _totalFolderCount =
                    0;

                _totalMailCount =
                    0;

                _processedMailCount =
                    0;

                ProgressPanel.Visibility =
                    Visibility.Visible;

                SetProgress(
                    0,
                    "Phase 1 von 2 – Outlook-Bestand wird gezählt");

                ProgressCountText.Text =
                    "0 E-Mails gefunden";

                ProgressFolderText.Text =
                    "";

                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType ==
                    null)
                {
                    throw new InvalidOperationException(
                        "Das klassische Microsoft Outlook wurde nicht gefunden.");
                }

                outlookApplication =
                    Activator.CreateInstance(
                        outlookType);

                dynamic outlook =
                    outlookApplication!;

                outlookNamespace =
                    outlook.GetNamespace(
                        "MAPI");

                dynamic outlookNs =
                    outlookNamespace!;

                stores =
                    outlookNs.Stores;

                dynamic outlookStores =
                    stores!;

                int storeCount =
                    outlookStores.Count;

                List<string> storeNames =
                    new List<string>();

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject =
                        null;

                    object? rootFolderObject =
                        null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(
                                storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(
                            storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

                        storeNames.Add(
                            storeName);

                        ProgressFolderText.Text =
                            $"Postfach {storeIndex} von {storeCount}: {storeName}";

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject !=
                            null)
                        {
                            await CountFolderAsync(
                                rootFolderObject,
                                storeName);
                        }

                        SetProgress(
                            ((double)storeIndex /
                             Math.Max(
                                 1,
                                 storeCount)) *
                            15.0,
                            $"Phase 1 von 2 – Bestand wird gezählt ({storeIndex}/{storeCount} Postfächer)");
                    }
                    finally
                    {
                        ReleaseComObject(
                            rootFolderObject);

                        ReleaseComObject(
                            storeObject);
                    }
                }

                if (_totalMailCount <=
                    0)
                {
                    throw new InvalidOperationException(
                        "Es wurden keine Outlook-E-Mails gefunden.");
                }

                Dictionary<string, IndexRecord> records =
                    new Dictionary<string, IndexRecord>(
                        StringComparer.OrdinalIgnoreCase);

                SetProgress(
                    15,
                    "Phase 2 von 2 – E-Mails und Anhänge werden indiziert");

                ProgressCountText.Text =
                    $"0 von {_totalMailCount:N0} E-Mails verarbeitet";

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject =
                        null;

                    object? rootFolderObject =
                        null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(
                                storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(
                            storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject !=
                            null)
                        {
                            await BuildFullFolderIndexAsync(
                                rootFolderObject,
                                storeName,
                                "",
                                records);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(
                            rootFolderObject);

                        ReleaseComObject(
                            storeObject);
                    }
                }

                SetProgress(
                    96,
                    "Indexdatei wird sicher gespeichert");

                ProgressCountText.Text =
                    $"{records.Count:N0} E-Mails wurden verarbeitet";

                ProgressFolderText.Text =
                    "Lokaler Suchindex wird abgeschlossen ...";

                await RefreshUi();

                DateTime syncTime =
                    DateTime.Now;

                await WriteIndexRecordsAsync(
                    records.Values,
                    storeCount,
                    syncTime);

                WriteLastSync(
                    syncTime);

                MailboxNamesText.Text =
                    "Verbunden: " +
                    string.Join(
                        "  •  ",
                        storeNames);

                SetProgress(
                    100,
                    "Fertig – E-Mails und Anhänge wurden vollständig indiziert");

                ProgressCountText.Text =
                    $"{records.Count:N0} E-Mails erfolgreich indiziert";

                ProgressFolderText.Text =
                    "Der neue Suchindex wurde erfolgreich gespeichert.";

                IndexStatusText.Text =
                    "Outlook-Inhaltsindex einschließlich Anhänge erfolgreich erstellt";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            0,
                            120,
                            70));

                IndexDetailsText.Text =
                    $"{storeCount} Postfächer | " +
                    $"{_totalFolderCount:N0} Ordner | " +
                    $"{records.Count:N0} E-Mails";

                SearchButton.IsEnabled =
                    true;

                CreateIndexButton.Content =
                    "Jetzt aktualisieren";

                LoadMailboxesFromIndex();
                UpdateLastSyncDisplay();

                SearchStatusText.Text =
                    "Index bereit. Betreff, Mailtext und Anhänge können durchsucht werden.";

                await Task.Delay(
                    1500);

                ProgressPanel.Visibility =
                    Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                SetFailedProgress(
                    "Indizierung abgebrochen",
                    ex.Message);

                IndexStatusText.Text =
                    "Outlook-Indizierung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(
                            180,
                            40,
                            40));

                IndexDetailsText.Text =
                    ex.Message;

                SearchButton.IsEnabled =
                    IsSearchableIndex();
            }
            finally
            {
                ReleaseComObject(
                    stores);

                ReleaseComObject(
                    outlookNamespace);

                ReleaseComObject(
                    outlookApplication);

                CreateIndexButton.IsEnabled =
                    true;

                ConnectOutlookButton.IsEnabled =
                    true;

                _updateInProgress =
                    false;
            }
        }

        private async Task CountFolderAsync(
            object folderObject,
            string storeName)
        {
            object? itemsObject =
                null;

            object? foldersObject =
                null;

            try
            {
                dynamic folder =
                    folderObject;

                _totalFolderCount++;

                string folderName =
                    SafeDynamicString(
                        folder,
                        "Name");

                if (string.IsNullOrWhiteSpace(
                    folderName))
                {
                    folderName =
                        "Unbekannter Ordner";
                }

                ProgressFolderText.Text =
                    $"{storeName} > {folderName}";

                try
                {
                    itemsObject =
                        folder.Items;

                    if (itemsObject !=
                        null)
                    {
                        dynamic items =
                            itemsObject;

                        int itemCount =
                            items.Count;

                        for (int i = 1;
                             i <= itemCount;
                             i++)
                        {
                            object? itemObject =
                                null;

                            try
                            {
                                itemObject =
                                    items.Item(i);

                                if (itemObject ==
                                    null)
                                {
                                    continue;
                                }

                                dynamic item =
                                    itemObject;

                                int itemClass =
                                    item.Class;

                                if (itemClass ==
                                    MailItemClass)
                                {
                                    _totalMailCount++;

                                    if (_totalMailCount %
                                        250 ==
                                        0)
                                    {
                                        ProgressCountText.Text =
                                            $"{_totalMailCount:N0} E-Mails gefunden";

                                        await RefreshUi();
                                    }
                                }
                            }
                            catch
                            {
                            }
                            finally
                            {
                                ReleaseComObject(
                                    itemObject);
                            }
                        }
                    }
                }
                catch
                {
                }

                foldersObject =
                    folder.Folders;

                if (foldersObject !=
                    null)
                {
                    dynamic folders =
                        foldersObject;

                    int subFolderCount =
                        folders.Count;

                    for (int i = 1;
                         i <= subFolderCount;
                         i++)
                    {
                        object? subFolderObject =
                            null;

                        try
                        {
                            subFolderObject =
                                folders.Item(i);

                            if (subFolderObject !=
                                null)
                            {
                                await CountFolderAsync(
                                    subFolderObject,
                                    storeName);
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            ReleaseComObject(
                                subFolderObject);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(
                    itemsObject);

                ReleaseComObject(
                    foldersObject);
            }
        }

        private async Task BuildFullFolderIndexAsync(
            object folderObject,
            string storeName,
            string parentPath,
            Dictionary<string, IndexRecord> records)
        {
            object? itemsObject =
                null;

            object? foldersObject =
                null;

            try
            {
                dynamic folder =
                    folderObject;

                string folderName =
                    SafeDynamicString(
                        folder,
                        "Name");

                if (string.IsNullOrWhiteSpace(
                    folderName))
                {
                    folderName =
                        "Unbekannter Ordner";
                }

                string currentPath =
                    string.IsNullOrWhiteSpace(
                        parentPath)
                        ? folderName
                        : parentPath +
                          " > " +
                          folderName;

                ProgressFolderText.Text =
                    $"{storeName} > {currentPath}";

                await RefreshUi();

                try
                {
                    itemsObject =
                        folder.Items;

                    dynamic items =
                        itemsObject!;

                    int itemCount =
                        items.Count;

                    for (int i = 1;
                         i <= itemCount;
                         i++)
                    {
                        object? itemObject =
                            null;

                        try
                        {
                            itemObject =
                                items.Item(i);

                            if (itemObject ==
                                null)
                            {
                                continue;
                            }

                            dynamic item =
                                itemObject;

                            int itemClass =
                                item.Class;

                            if (itemClass !=
                                MailItemClass)
                            {
                                continue;
                            }

                            IndexRecord record =
                                BuildIndexRecord(
                                    item,
                                    storeName,
                                    currentPath);

                            if (!string.IsNullOrWhiteSpace(
                                record.EntryId))
                            {
                                records[
                                    record.EntryId] =
                                    record;
                            }

                            _processedMailCount++;

                            if (_processedMailCount %
                                10 ==
                                0)
                            {
                                int displayedProcessed =
                                    Math.Min(
                                        _processedMailCount,
                                        _totalMailCount);

                                double mailRatio =
                                    _totalMailCount >
                                    0
                                        ? (double)displayedProcessed /
                                          _totalMailCount
                                        : 0;

                                double percent =
                                    15.0 +
                                    mailRatio *
                                    80.0;

                                SetProgress(
                                    percent,
                                    $"Phase 2 von 2 – {displayedProcessed:N0} von {_totalMailCount:N0} E-Mails");

                                ProgressCountText.Text =
                                    $"{displayedProcessed:N0} von {_totalMailCount:N0} E-Mails verarbeitet";

                                await RefreshUi();
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            ReleaseComObject(
                                itemObject);
                        }
                    }
                }
                catch
                {
                }

                foldersObject =
                    folder.Folders;

                if (foldersObject !=
                    null)
                {
                    dynamic folders =
                        foldersObject;

                    int subFolderCount =
                        folders.Count;

                    for (int i = 1;
                         i <= subFolderCount;
                         i++)
                    {
                        object? subFolderObject =
                            null;

                        try
                        {
                            subFolderObject =
                                folders.Item(i);

                            if (subFolderObject !=
                                null)
                            {
                                await BuildFullFolderIndexAsync(
                                    subFolderObject,
                                    storeName,
                                    currentPath,
                                    records);
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            ReleaseComObject(
                                subFolderObject);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(
                    itemsObject);

                ReleaseComObject(
                    foldersObject);
            }
        }

        // =========================================================
        // INDEX LADEN / SPEICHERN
        // =========================================================

        private Dictionary<string, IndexRecord>
            LoadExistingIndexRecords()
        {
            Dictionary<string, IndexRecord> records =
                new Dictionary<string, IndexRecord>(
                    StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(
                _indexPath))
            {
                return records;
            }

            using StreamReader reader =
                new StreamReader(
                    _indexPath,
                    Encoding.UTF8,
                    true);

            string? line;

            int legacyCounter =
                0;

            while ((line =
                reader.ReadLine()) != null)
            {
                string[] columns =
                    line.Split('\t');

                if (columns.Length <
                    14)
                {
                    continue;
                }

                if (!int.TryParse(
                    columns[0],
                    out _))
                {
                    continue;
                }

                IndexRecord record =
                    new IndexRecord
                    {
                        Date =
                            columns[1],

                        Sender =
                            columns[2],

                        Recipient =
                            columns[3],

                        Cc =
                            columns[4],

                        Mailbox =
                            columns[5],

                        Subject =
                            columns[6],

                        Folder =
                            columns[7],

                        Flag =
                            columns[8],

                        Attachment =
                            columns[9],

                        Categories =
                            columns[10],

                        ConversationId =
                            columns[11],

                        EntryId =
                            columns[12],

                        Body =
                            columns[13],

                        AttachmentNames =
                            columns.Length >=
                            15
                                ? columns[14]
                                : "",

                        AttachmentText =
                            columns.Length >=
                            16
                                ? columns[15]
                                : ""
                    };

                string key =
                    record.EntryId;

                if (string.IsNullOrWhiteSpace(
                    key))
                {
                    legacyCounter++;

                    key =
                        "__LEGACY__" +
                        legacyCounter;
                }

                records[
                    key] =
                    record;
            }

            return records;
        }

        private async Task WriteIndexRecordsAsync(
            IEnumerable<IndexRecord> records,
            int storeCount,
            DateTime syncTime)
        {
            Directory.CreateDirectory(
                _indexFolder);

            string temporaryPath =
                _indexPath +
                ".tmp";

            List<IndexRecord> orderedRecords =
                records
                    .OrderByDescending(
                        x =>
                            ParseIndexDate(
                                x.Date)
                            ??
                            DateTime.MinValue)
                    .ToList();

            try
            {
                if (File.Exists(
                    temporaryPath))
                {
                    File.Delete(
                        temporaryPath);
                }
            }
            catch
            {
                temporaryPath =
                    _indexPath +
                    "." +
                    Guid.NewGuid()
                        .ToString(
                            "N") +
                    ".tmp";
            }

            await using (
                FileStream stream =
                    new FileStream(
                        temporaryPath,
                        FileMode.Create,
                        FileAccess.Write,
                        FileShare.None,
                        65536,
                        true))
            {
                await using (
                    StreamWriter writer =
                        new StreamWriter(
                            stream,
                            new UTF8Encoding(
                                true)))
                {
                    await writer.WriteLineAsync(
                        "Datenfinder Gemeinde Reichenau PRO");

                    await writer.WriteLineAsync(
                        "Outlook-Inhaltsindex");

                    await writer.WriteLineAsync(
                        $"Schema: {IndexSchema}");

                    await writer.WriteLineAsync(
                        $"Letzte Aktualisierung: {syncTime:dd.MM.yyyy HH:mm:ss}");

                    await writer.WriteLineAsync(
                        $"Postfächer: {storeCount}");

                    await writer.WriteLineAsync(
                        $"E-Mails: {orderedRecords.Count}");

                    await writer.WriteLineAsync();

                    await writer.WriteLineAsync(
                        "Nr.\t" +
                        "Datum\t" +
                        "Absender\t" +
                        "Empfänger\t" +
                        "CC\t" +
                        "Postfach\t" +
                        "Betreff\t" +
                        "Ordner\t" +
                        "Kennzeichnung\t" +
                        "Anhang\t" +
                        "Kategorien\t" +
                        "ConversationID\t" +
                        "EntryID\t" +
                        "E-Mail-Text\t" +
                        "Anhang-Namen\t" +
                        "Anhang-Texte");

                    int number =
                        0;

                    foreach (
                        IndexRecord record
                        in orderedRecords)
                    {
                        number++;

                        await writer.WriteLineAsync(
                            SerializeIndexRecord(
                                number,
                                record));
                    }

                    await writer.FlushAsync();
                }
            }

            File.Move(
                temporaryPath,
                _indexPath,
                true);
        }

        private static string SerializeIndexRecord(
            int number,
            IndexRecord record)
        {
            return
                $"{number}\t" +
                $"{CleanIndexText(record.Date)}\t" +
                $"{CleanIndexText(record.Sender)}\t" +
                $"{CleanIndexText(record.Recipient)}\t" +
                $"{CleanIndexText(record.Cc)}\t" +
                $"{CleanIndexText(record.Mailbox)}\t" +
                $"{CleanIndexText(record.Subject)}\t" +
                $"{CleanIndexText(record.Folder)}\t" +
                $"{CleanIndexText(record.Flag)}\t" +
                $"{CleanIndexText(record.Attachment)}\t" +
                $"{CleanIndexText(record.Categories)}\t" +
                $"{CleanIndexText(record.ConversationId)}\t" +
                $"{CleanIndexText(record.EntryId)}\t" +
                $"{CleanIndexText(record.Body)}\t" +
                $"{CleanIndexText(record.AttachmentNames)}\t" +
                $"{CleanIndexText(record.AttachmentText)}";
        }

        // =========================================================
        // E-MAIL + ANHÄNGE EINLESEN
        // =========================================================

        private static IndexRecord BuildIndexRecord(
            dynamic item,
            string storeName,
            string currentPath)
        {
            AttachmentIndexData attachmentData =
                ExtractAttachmentData(
                    item);

            return new IndexRecord
            {
                Date =
                    SafeDynamicDateTime(
                        item),

                Sender =
                    SafeDynamicString(
                        item,
                        "SenderName"),

                Recipient =
                    SafeDynamicString(
                        item,
                        "To"),

                Cc =
                    SafeDynamicString(
                        item,
                        "CC"),

                Mailbox =
                    storeName,

                Subject =
                    SafeDynamicString(
                        item,
                        "Subject"),

                Folder =
                    currentPath,

                Flag =
                    GetFlagDescription(
                        item),

                Attachment =
                    attachmentData.HasAttachments
                        ? "1"
                        : "0",

                Categories =
                    SafeDynamicString(
                        item,
                        "Categories"),

                ConversationId =
                    SafeDynamicString(
                        item,
                        "ConversationID"),

                EntryId =
                    SafeDynamicString(
                        item,
                        "EntryID"),

                Body =
                    SafeDynamicString(
                        item,
                        "Body"),

                AttachmentNames =
                    attachmentData.Names,

                AttachmentText =
                    attachmentData.SearchText
            };
        }

        private static AttachmentIndexData
            ExtractAttachmentData(
                dynamic item)
        {
            object? attachmentsObject =
                null;

            List<string> names =
                new List<string>();

            StringBuilder text =
                new StringBuilder();

            int totalCharacters =
                0;

            try
            {
                attachmentsObject =
                    item.Attachments;

                if (attachmentsObject ==
                    null)
                {
                    return new AttachmentIndexData();
                }

                dynamic attachments =
                    attachmentsObject;

                int count =
                    attachments.Count;

                for (int i = 1;
                     i <= count;
                     i++)
                {
                    object? attachmentObject =
                        null;

                    string? temporaryPath =
                        null;

                    try
                    {
                        attachmentObject =
                            attachments.Item(
                                i);

                        dynamic attachment =
                            attachmentObject;

                        string fileName =
                            "";

                        try
                        {
                            fileName =
                                attachment.FileName ??
                                "";
                        }
                        catch
                        {
                        }

                        if (string.IsNullOrWhiteSpace(
                            fileName))
                        {
                            fileName =
                                $"Anhang {i}";
                        }

                        names.Add(
                            fileName);

                        string extension =
                            Path.GetExtension(
                                fileName)
                            .ToLowerInvariant();

                        if (!IsSupportedAttachment(
                            extension))
                        {
                            continue;
                        }

                        long size =
                            0;

                        try
                        {
                            size =
                                Convert.ToInt64(
                                    attachment.Size);
                        }
                        catch
                        {
                        }

                        if (size >
                            MaximumAttachmentBytes)
                        {
                            continue;
                        }

                        string tempFolder =
                            Path.Combine(
                                Path.GetTempPath(),
                                "DatenfinderGemeindeReichenau");

                        Directory.CreateDirectory(
                            tempFolder);

                        temporaryPath =
                            Path.Combine(
                                tempFolder,
                                Guid.NewGuid()
                                    .ToString(
                                        "N") +
                                extension);

                        attachment.SaveAsFile(
                            temporaryPath);

                        string extracted =
                            ExtractTextFromAttachment(
                                temporaryPath,
                                extension);

                        if (string.IsNullOrWhiteSpace(
                            extracted))
                        {
                            continue;
                        }

                        extracted =
                            NormalizeExtractedText(
                                extracted);

                        if (extracted.Length >
                            MaximumCharactersPerAttachment)
                        {
                            extracted =
                                extracted.Substring(
                                    0,
                                    MaximumCharactersPerAttachment);
                        }

                        int remaining =
                            MaximumAttachmentCharactersPerMail -
                            totalCharacters;

                        if (remaining <=
                            0)
                        {
                            break;
                        }

                        if (extracted.Length >
                            remaining)
                        {
                            extracted =
                                extracted.Substring(
                                    0,
                                    remaining);
                        }

                        string markerName =
                            fileName
                                .Replace(
                                    "]]]",
                                    "")
                                .Replace(
                                    "[[[",
                                    "");

                        text.Append(
                            "[[[ATTACHMENT:");

                        text.Append(
                            markerName);

                        text.Append(
                            "]]] ");

                        text.Append(
                            extracted);

                        text.Append(
                            " [[[/ATTACHMENT]]] ");

                        totalCharacters +=
                            extracted.Length;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (!string.IsNullOrWhiteSpace(
                            temporaryPath))
                        {
                            try
                            {
                                if (File.Exists(
                                    temporaryPath))
                                {
                                    File.Delete(
                                        temporaryPath);
                                }
                            }
                            catch
                            {
                            }
                        }

                        ReleaseComObject(
                            attachmentObject);
                    }
                }

                return new AttachmentIndexData
                {
                    HasAttachments =
                        count > 0,

                    Names =
                        string.Join(
                            " | ",
                            names),

                    SearchText =
                        text.ToString()
                };
            }
            catch
            {
                return new AttachmentIndexData
                {
                    HasAttachments =
                        names.Count >
                        0,

                    Names =
                        string.Join(
                            " | ",
                            names),

                    SearchText =
                        text.ToString()
                };
            }
            finally
            {
                ReleaseComObject(
                    attachmentsObject);
            }
        }

        private static bool IsSupportedAttachment(
            string extension)
        {
            return
                extension == ".pdf" ||
                extension == ".docx" ||
                extension == ".xlsx" ||
                extension == ".txt" ||
                extension == ".csv";
        }

        private static string ExtractTextFromAttachment(
            string filePath,
            string extension)
        {
            try
            {
                return extension switch
                {
                    ".pdf" =>
                        ExtractPdfText(
                            filePath),

                    ".docx" =>
                        ExtractDocxText(
                            filePath),

                    ".xlsx" =>
                        ExtractXlsxText(
                            filePath),

                    ".txt" =>
                        File.ReadAllText(
                            filePath),

                    ".csv" =>
                        File.ReadAllText(
                            filePath),

                    _ =>
                        ""
                };
            }
            catch
            {
                return "";
            }
        }

        private static string ExtractPdfText(
            string filePath)
        {
            StringBuilder result =
                new StringBuilder();

            using PdfDocument document =
                PdfDocument.Open(
                    filePath);

            foreach (var page
                in document.GetPages())
            {
                if (!string.IsNullOrWhiteSpace(
                    page.Text))
                {
                    result.AppendLine(
                        page.Text);
                }
            }

            return
                result.ToString();
        }

        private static string ExtractDocxText(
            string filePath)
        {
            using ZipArchive archive =
                ZipFile.OpenRead(
                    filePath);

            ZipArchiveEntry? documentEntry =
                archive.GetEntry(
                    "word/document.xml");

            if (documentEntry ==
                null)
            {
                return "";
            }

            using Stream stream =
                documentEntry.Open();

            XDocument document =
                XDocument.Load(
                    stream);

            IEnumerable<string> textNodes =
                document
                    .Descendants()
                    .Where(
                        element =>
                            element.Name.LocalName ==
                            "t")
                    .Select(
                        element =>
                            element.Value);

            return
                string.Join(
                    " ",
                    textNodes);
        }

        private static string ExtractXlsxText(
            string filePath)
        {
            using ZipArchive archive =
                ZipFile.OpenRead(
                    filePath);

            List<string> sharedStrings =
                new List<string>();

            ZipArchiveEntry? sharedEntry =
                archive.GetEntry(
                    "xl/sharedStrings.xml");

            if (sharedEntry !=
                null)
            {
                using Stream sharedStream =
                    sharedEntry.Open();

                XDocument sharedDocument =
                    XDocument.Load(
                        sharedStream);

                foreach (XElement item
                    in sharedDocument
                        .Descendants()
                        .Where(
                            element =>
                                element.Name.LocalName ==
                                "si"))
                {
                    string text =
                        string.Join(
                            "",
                            item
                                .Descendants()
                                .Where(
                                    element =>
                                        element.Name.LocalName ==
                                        "t")
                                .Select(
                                    element =>
                                        element.Value));

                    sharedStrings.Add(
                        text);
                }
            }

            StringBuilder result =
                new StringBuilder();

            IEnumerable<ZipArchiveEntry> worksheets =
                archive.Entries
                    .Where(
                        entry =>
                            entry.FullName.StartsWith(
                                "xl/worksheets/sheet",
                                StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(
                                ".xml",
                                StringComparison.OrdinalIgnoreCase));

            foreach (
                ZipArchiveEntry worksheet
                in worksheets)
            {
                using Stream worksheetStream =
                    worksheet.Open();

                XDocument sheetDocument =
                    XDocument.Load(
                        worksheetStream);

                foreach (XElement cell
                    in sheetDocument
                        .Descendants()
                        .Where(
                            element =>
                                element.Name.LocalName ==
                                "c"))
                {
                    string type =
                        cell.Attribute(
                            "t")
                            ?.Value ??
                        "";

                    XElement? valueElement =
                        cell.Elements()
                            .FirstOrDefault(
                                element =>
                                    element.Name.LocalName ==
                                    "v");

                    if (type ==
                        "inlineStr")
                    {
                        string inlineText =
                            string.Join(
                                "",
                                cell
                                    .Descendants()
                                    .Where(
                                        element =>
                                            element.Name.LocalName ==
                                            "t")
                                    .Select(
                                        element =>
                                            element.Value));

                        if (!string.IsNullOrWhiteSpace(
                            inlineText))
                        {
                            result.Append(
                                inlineText);

                            result.Append(
                                ' ');
                        }

                        continue;
                    }

                    if (valueElement ==
                        null)
                    {
                        continue;
                    }

                    string rawValue =
                        valueElement.Value;

                    if (type ==
                            "s" &&
                        int.TryParse(
                            rawValue,
                            out int sharedIndex) &&
                        sharedIndex >=
                            0 &&
                        sharedIndex <
                            sharedStrings.Count)
                    {
                        result.Append(
                            sharedStrings[
                                sharedIndex]);
                    }
                    else
                    {
                        result.Append(
                            rawValue);
                    }

                    result.Append(
                        ' ');
                }
            }

            return
                result.ToString();
        }

        private static string NormalizeExtractedText(
            string text)
        {
            if (string.IsNullOrWhiteSpace(
                text))
            {
                return "";
            }

            return text
                .Replace(
                    "\r",
                    " ")
                .Replace(
                    "\n",
                    " ")
                .Replace(
                    "\t",
                    " ")
                .Trim();
        }

        private static DateTime?
            SafeLastModificationTime(
                dynamic item)
        {
            try
            {
                return
                    item.LastModificationTime;
            }
            catch
            {
                try
                {
                    return
                        item.ReceivedTime;
                }
                catch
                {
                    try
                    {
                        return
                            item.SentOn;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }

        // =========================================================
        // FORTSCHRITT
        // =========================================================

        private void SetProgress(
            double percent,
            string phaseText)
        {
            double safePercent =
                Math.Max(
                    0,
                    Math.Min(
                        100,
                        percent));

            IndexProgressBar.IsIndeterminate =
                false;

            IndexProgressBar.Value =
                safePercent;

            ProgressPercentText.Text =
                $"{safePercent:0} %";

            ProgressPhaseText.Text =
                phaseText;
        }

        private void SetFailedProgress(
            string phaseText,
            string errorText)
        {
            IndexProgressBar.IsIndeterminate =
                false;

            IndexProgressBar.Minimum =
                0;

            IndexProgressBar.Maximum =
                100;

            IndexProgressBar.Value =
                0;

            ProgressPercentText.Text =
                "0 %";

            ProgressPhaseText.Text =
                phaseText;

            ProgressCountText.Text =
                "Vorgang wurde nicht abgeschlossen.";

            ProgressFolderText.Text =
                errorText;
        }

        // =========================================================
        // SUCHE
        // =========================================================

        private async void SearchButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            await ExecuteSearchAsync();
        }

        private async void SearchTextBox_KeyDown(
            object sender,
            KeyEventArgs e)
        {
            if (e.Key ==
                    Key.Enter &&
                SearchButton.IsEnabled)
            {
                await ExecuteSearchAsync();
            }
        }

        private void ResetFiltersButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            FromDatePicker.SelectedDate =
                null;

            ToDatePicker.SelectedDate =
                null;

            MailboxComboBox.SelectedIndex =
                0;

            AttachmentComboBox.SelectedIndex =
                0;

            FlagComboBox.SelectedIndex =
                0;

            SortComboBox.SelectedIndex =
                0;

            SenderFilterTextBox.Text =
                "";

            RecipientFilterTextBox.Text =
                "";

            SearchSubjectCheckBox.IsChecked =
                true;

            SearchBodyCheckBox.IsChecked =
                true;

            SearchAttachmentsCheckBox.IsChecked =
                false;

            ActiveFiltersText.Text =
                "Aktive Filter: keine";

            SearchTextBox.Focus();
        }

        private async Task ExecuteSearchAsync()
        {
            if (!IsSearchableIndex())
            {
                SearchStatusText.Text =
                    "Es wurde noch kein durchsuchbarer Index gefunden.";

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                return;
            }

            string query =
                SearchTextBox.Text.Trim();

            string senderFilter =
                SenderFilterTextBox.Text.Trim();

            string recipientFilter =
                RecipientFilterTextBox.Text.Trim();

            DateTime? fromDate =
                FromDatePicker.SelectedDate;

            DateTime? toDate =
                ToDatePicker.SelectedDate;

            if (fromDate.HasValue &&
                toDate.HasValue &&
                fromDate.Value.Date >
                toDate.Value.Date)
            {
                SearchStatusText.Text =
                    "Das Von-Datum liegt nach dem Bis-Datum.";

                return;
            }

            bool searchSubject =
                SearchSubjectCheckBox.IsChecked ==
                true;

            bool searchBody =
                SearchBodyCheckBox.IsChecked ==
                true;

            bool searchAttachments =
                SearchAttachmentsCheckBox.IsChecked ==
                true;

            if (!string.IsNullOrWhiteSpace(
                    query) &&
                !searchSubject &&
                !searchBody &&
                !searchAttachments)
            {
                SearchStatusText.Text =
                    "Bitte mindestens einen Suchbereich auswählen: Betreff, Mailtext oder Anhänge.";

                return;
            }

            if (searchAttachments &&
                !HasAttachmentIndex())
            {
                SearchStatusText.Text =
                    "Die Anhangsuche benötigt den neuen Index. Bitte zuerst den Anhangindex neu aufbauen.";

                return;
            }

            string mailbox =
                MailboxComboBox
                    .SelectedItem
                    ?.ToString()
                ??
                "Alle Postfächer";

            string attachment =
                GetComboBoxText(
                    AttachmentComboBox);

            string flag =
                GetComboBoxText(
                    FlagComboBox);

            string sort =
                GetComboBoxText(
                    SortComboBox);

            bool noCriteria =
                string.IsNullOrWhiteSpace(
                    query) &&
                string.IsNullOrWhiteSpace(
                    senderFilter) &&
                string.IsNullOrWhiteSpace(
                    recipientFilter) &&
                !fromDate.HasValue &&
                !toDate.HasValue &&
                mailbox ==
                    "Alle Postfächer" &&
                attachment ==
                    "Alle" &&
                flag ==
                    "Alle";

            if (noCriteria)
            {
                ActiveFiltersText.Text =
                    "Aktive Filter: keine";

                SearchStatusText.Text =
                    "Bitte einen Suchbegriff eingeben oder mindestens einen Filter auswählen.";

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                return;
            }

            SearchButton.IsEnabled =
                false;

            SearchTextBox.IsEnabled =
                false;

            SearchStatusText.Text =
                "Index wird durchsucht ...";

            try
            {
                SearchOptions options =
                    new SearchOptions
                    {
                        Query =
                            query,

                        FromDate =
                            fromDate,

                        ToDate =
                            toDate,

                        Mailbox =
                            mailbox,

                        Attachment =
                            attachment,

                        Flag =
                            flag,

                        SenderFilter =
                            senderFilter,

                        RecipientFilter =
                            recipientFilter,

                        SearchSubject =
                            searchSubject,

                        SearchBody =
                            searchBody,

                        SearchAttachments =
                            searchAttachments,

                        Sort =
                            sort
                    };

                ActiveFiltersText.Text =
                    BuildActiveFiltersText(
                        options);

                SearchResponse response =
                    await Task.Run(
                        () =>
                            SearchIndex(
                                options));

                SearchResultsGrid.ItemsSource =
                    response.Results;

                SearchResultsGrid.SelectedItems.Clear();

                UpdateSelectionDisplay();

                if (response.Results.Count ==
                    0)
                {
                    SearchStatusText.Text =
                        "Keine passenden E-Mails gefunden.";

                    SearchResultsGrid.Visibility =
                        Visibility.Collapsed;
                }
                else
                {
                    string limitText =
                        response.WasLimited
                            ? $" – angezeigt werden die ersten {MaximumSearchResults:N0}"
                            : "";

                    SearchStatusText.Text =
                        $"{response.TotalMatches:N0} Treffer gefunden{limitText}.";

                    SearchResultsGrid.Visibility =
                        Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                SearchStatusText.Text =
                    $"Suche fehlgeschlagen: {ex.Message}";

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;
            }
            finally
            {
                SearchButton.IsEnabled =
                    true;

                SearchTextBox.IsEnabled =
                    true;

                SearchTextBox.Focus();
            }
        }

        private static string BuildActiveFiltersText(
            SearchOptions options)
        {
            List<string> filters =
                new List<string>();

            if (options.FromDate.HasValue ||
                options.ToDate.HasValue)
            {
                string from =
                    options.FromDate.HasValue
                        ? options.FromDate.Value
                            .ToString(
                                "dd.MM.yyyy")
                        : "offen";

                string to =
                    options.ToDate.HasValue
                        ? options.ToDate.Value
                            .ToString(
                                "dd.MM.yyyy")
                        : "offen";

                filters.Add(
                    $"Zeitraum {from}–{to}");
            }

            if (options.Mailbox !=
                "Alle Postfächer")
            {
                filters.Add(
                    $"Postfach: {options.Mailbox}");
            }

            if (!string.IsNullOrWhiteSpace(
                options.SenderFilter))
            {
                filters.Add(
                    $"Absender: {options.SenderFilter}");
            }

            if (!string.IsNullOrWhiteSpace(
                options.RecipientFilter))
            {
                filters.Add(
                    $"Empfänger: {options.RecipientFilter}");
            }

            if (options.Attachment !=
                "Alle")
            {
                filters.Add(
                    options.Attachment);
            }

            if (options.Flag !=
                "Alle")
            {
                filters.Add(
                    $"Kennzeichnung: {options.Flag}");
            }

            List<string> areas =
                new List<string>();

            if (options.SearchSubject)
            {
                areas.Add(
                    "Betreff");
            }

            if (options.SearchBody)
            {
                areas.Add(
                    "Mailtext");
            }

            if (options.SearchAttachments)
            {
                areas.Add(
                    "Anhänge");
            }

            filters.Add(
                "Suche in: " +
                string.Join(
                    ", ",
                    areas));

            return
                "Aktive Filter: " +
                string.Join(
                    "  •  ",
                    filters);
        }

        private SearchResponse SearchIndex(
            SearchOptions options)
        {
            List<SearchResult> matches =
                new List<SearchResult>();

            string[] searchWords =
                options.Query.Split(
                    new[] { ' ' },
                    StringSplitOptions
                        .RemoveEmptyEntries |
                    StringSplitOptions
                        .TrimEntries);

            foreach (
                IndexRecord record
                in ReadIndexRecords())
            {
                DateTime? mailDate =
                    ParseIndexDate(
                        record.Date);

                if (options.FromDate.HasValue &&
                    (!mailDate.HasValue ||
                     mailDate.Value.Date <
                     options.FromDate.Value.Date))
                {
                    continue;
                }

                if (options.ToDate.HasValue &&
                    (!mailDate.HasValue ||
                     mailDate.Value.Date >
                     options.ToDate.Value.Date))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(
                        options.SenderFilter) &&
                    !record.Sender.Contains(
                        options.SenderFilter,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(
                        options.RecipientFilter) &&
                    !record.Recipient.Contains(
                        options.RecipientFilter,
                        StringComparison.OrdinalIgnoreCase) &&
                    !record.Cc.Contains(
                        options.RecipientFilter,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (options.Mailbox !=
                        "Alle Postfächer" &&
                    !record.Mailbox.Equals(
                        options.Mailbox,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                bool hasAttachment =
                    record.Attachment ==
                    "1";

                if (options.Attachment ==
                        "Mit Anhang" &&
                    !hasAttachment)
                {
                    continue;
                }

                if (options.Attachment ==
                        "Ohne Anhang" &&
                    hasAttachment)
                {
                    continue;
                }

                bool isFlagged =
                    !record.Flag.Equals(
                        "Keine Kennzeichnung",
                        StringComparison.OrdinalIgnoreCase);

                if (options.Flag ==
                        "Gekennzeichnet" &&
                    !isFlagged)
                {
                    continue;
                }

                if (options.Flag ==
                        "Nicht gekennzeichnet" &&
                    isFlagged)
                {
                    continue;
                }

                if (options.Flag ==
                        "Erledigt / Häkchen" &&
                    !record.Flag.Contains(
                        "Erledigt",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (options.Flag ==
                        "Fähnchen" &&
                    !record.Flag.Contains(
                        "Fähnchen",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string foundIn =
                    "Filter";

                if (searchWords.Length >
                    0)
                {
                    StringBuilder searchable =
                        new StringBuilder();

                    if (options.SearchSubject)
                    {
                        searchable.Append(
                            record.Subject);

                        searchable.Append(
                            ' ');
                    }

                    if (options.SearchBody)
                    {
                        searchable.Append(
                            record.Body);

                        searchable.Append(
                            ' ');
                    }

                    if (options.SearchAttachments)
                    {
                        searchable.Append(
                            record.AttachmentNames);

                        searchable.Append(
                            ' ');

                        searchable.Append(
                            record.AttachmentText);
                    }

                    if (!AllWordsFound(
                        searchable.ToString(),
                        searchWords))
                    {
                        continue;
                    }

                    List<string> locations =
                        new List<string>();

                    if (options.SearchSubject &&
                        AllWordsFound(
                            record.Subject,
                            searchWords))
                    {
                        locations.Add(
                            "Betreff");
                    }

                    if (options.SearchBody &&
                        AllWordsFound(
                            record.Body,
                            searchWords))
                    {
                        locations.Add(
                            "Mailtext");
                    }

                    if (options.SearchAttachments)
                    {
                        List<string> matchingAttachments =
                            FindMatchingAttachmentNames(
                                record.AttachmentText,
                                searchWords);

                        if (matchingAttachments.Count >
                            0)
                        {
                            foreach (
                                string fileName
                                in matchingAttachments)
                            {
                                locations.Add(
                                    "Anhang: " +
                                    fileName);
                            }
                        }
                        else if (
                            AllWordsFound(
                                record.AttachmentNames +
                                " " +
                                record.AttachmentText,
                                searchWords))
                        {
                            locations.Add(
                                "Anhang");
                        }
                    }

                    foundIn =
                        locations.Count >
                            0
                            ? string.Join(
                                " + ",
                                locations)
                            : "Mehrere Bereiche";
                }

                matches.Add(
                    ConvertRecordToSearchResult(
                        record,
                        foundIn));
            }

            IEnumerable<SearchResult> sorted =
                options.Sort ==
                    "Älteste zuerst"
                    ? matches.OrderBy(
                        x =>
                            x.SortDate)
                    : matches.OrderByDescending(
                        x =>
                            x.SortDate);

            int totalMatches =
                matches.Count;

            List<SearchResult> displayed =
                sorted
                    .Take(
                        MaximumSearchResults)
                    .ToList();

            return new SearchResponse
            {
                Results =
                    displayed,

                TotalMatches =
                    totalMatches,

                WasLimited =
                    totalMatches >
                    MaximumSearchResults
            };
        }

        private List<IndexRecord> ReadIndexRecords()
        {
            List<IndexRecord> records =
                new List<IndexRecord>();

            if (!File.Exists(
                _indexPath))
            {
                return records;
            }

            using StreamReader reader =
                new StreamReader(
                    _indexPath,
                    Encoding.UTF8,
                    true);

            string? line;

            while ((line =
                reader.ReadLine()) != null)
            {
                string[] columns =
                    line.Split(
                        '\t');

                if (columns.Length <
                    14)
                {
                    continue;
                }

                if (!int.TryParse(
                    columns[0],
                    out _))
                {
                    continue;
                }

                records.Add(
                    new IndexRecord
                    {
                        Date =
                            columns[1],

                        Sender =
                            columns[2],

                        Recipient =
                            columns[3],

                        Cc =
                            columns[4],

                        Mailbox =
                            columns[5],

                        Subject =
                            columns[6],

                        Folder =
                            columns[7],

                        Flag =
                            columns[8],

                        Attachment =
                            columns[9],

                        Categories =
                            columns[10],

                        ConversationId =
                            columns[11],

                        EntryId =
                            columns[12],

                        Body =
                            columns[13],

                        AttachmentNames =
                            columns.Length >=
                            15
                                ? columns[14]
                                : "",

                        AttachmentText =
                            columns.Length >=
                            16
                                ? columns[15]
                                : ""
                    });
            }

            return records;
        }

        private static SearchResult
            ConvertRecordToSearchResult(
                IndexRecord record,
                string foundIn =
                    "")
        {
            DateTime? date =
                ParseIndexDate(
                    record.Date);

            return new SearchResult
            {
                SortDate =
                    date ??
                    DateTime.MinValue,

                Date =
                    date.HasValue
                        ? date.Value.ToString(
                            "dd.MM.yyyy HH:mm")
                        : "",

                Mailbox =
                    record.Mailbox,

                Recipient =
                    record.Recipient,

                Cc =
                    record.Cc,

                Sender =
                    record.Sender,

                Flag =
                    record.Flag,

                Attachment =
                    record.Attachment ==
                        "1"
                        ? "Ja"
                        : "Nein",

                AttachmentNames =
                    string.IsNullOrWhiteSpace(
                        record.AttachmentNames)
                        ? "—"
                        : record.AttachmentNames,

                FoundIn =
                    foundIn,

                Subject =
                    record.Subject,

                Folder =
                    record.Folder,

                Categories =
                    record.Categories,

                ConversationId =
                    record.ConversationId,

                EntryId =
                    record.EntryId,

                Body =
                    record.Body
            };
        }

        private static bool AllWordsFound(
            string text,
            IEnumerable<string> words)
        {
            foreach (
                string word
                in words)
            {
                if (!text.Contains(
                    word,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private static List<string>
            FindMatchingAttachmentNames(
                string attachmentText,
                string[] searchWords)
        {
            List<string> results =
                new List<string>();

            const string startMarker =
                "[[[ATTACHMENT:";

            const string nameEndMarker =
                "]]]";

            const string endMarker =
                "[[[/ATTACHMENT]]]";

            int position =
                0;

            while (position <
                attachmentText.Length)
            {
                int start =
                    attachmentText.IndexOf(
                        startMarker,
                        position,
                        StringComparison.Ordinal);

                if (start <
                    0)
                {
                    break;
                }

                int nameStart =
                    start +
                    startMarker.Length;

                int nameEnd =
                    attachmentText.IndexOf(
                        nameEndMarker,
                        nameStart,
                        StringComparison.Ordinal);

                if (nameEnd <
                    0)
                {
                    break;
                }

                string fileName =
                    attachmentText.Substring(
                        nameStart,
                        nameEnd -
                        nameStart);

                int contentStart =
                    nameEnd +
                    nameEndMarker.Length;

                int end =
                    attachmentText.IndexOf(
                        endMarker,
                        contentStart,
                        StringComparison.Ordinal);

                if (end <
                    0)
                {
                    break;
                }

                string content =
                    attachmentText.Substring(
                        contentStart,
                        end -
                        contentStart);

                if (AllWordsFound(
                    fileName +
                    " " +
                    content,
                    searchWords))
                {
                    results.Add(
                        fileName);
                }

                position =
                    end +
                    endMarker.Length;
            }

            return results;
        }

        // =========================================================
        // MEHRFACHAUSWAHL
        // =========================================================

        private void SearchResultsGrid_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            UpdateSelectionDisplay();
        }

        private void UpdateSelectionDisplay()
        {
            int count = SearchResultsGrid.SelectedItems.Count;

            SelectionCountText.Text =
                count switch
                {
                    0 => "Keine Mail ausgewählt",
                    1 => "1 Mail ausgewählt",
                    _ => $"{count:N0} Mails ausgewählt"
                };
        }

        private void SelectAllButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (SearchResultsGrid.Items.Count == 0)
            {
                return;
            }

            SearchResultsGrid.SelectAll();
            UpdateSelectionDisplay();
        }

        private void ClearSelectionButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            SearchResultsGrid.UnselectAll();
            UpdateSelectionDisplay();
        }

        private List<SearchResult> GetSelectedSearchResults()
        {
            return SearchResultsGrid
                .SelectedItems
                .OfType<SearchResult>()
                .ToList();
        }

        // =========================================================
        // EXPORT
        // =========================================================

        private void ExportSelectedButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            ExportSelectedResults();
        }

        private void ExportSelectedMenuItem_Click(
            object sender,
            RoutedEventArgs e)
        {
            ExportSelectedResults();
        }

        private void ExportSelectedResults()
        {
            List<SearchResult> selected =
                GetSelectedSearchResults();

            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Bitte zuerst eine oder mehrere E-Mails auswählen.",
                    "E-Mails exportieren",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            ExportSearchResults(selected);
        }

        private void ExportSearchResults(
            IReadOnlyList<SearchResult> results)
        {
            Microsoft.Win32.OpenFolderDialog folderDialog =
                new Microsoft.Win32.OpenFolderDialog
                {
                    Title =
                        "Zielordner für den Mail-Export auswählen",

                    Multiselect =
                        false
                };

            bool? dialogResult =
                folderDialog.ShowDialog(this);

            if (dialogResult != true)
            {
                return;
            }

            string targetFolder =
                folderDialog.FolderName;

            if (string.IsNullOrWhiteSpace(targetFolder))
            {
                return;
            }

            Directory.CreateDirectory(targetFolder);

            int exported = 0;
            int failed = 0;

            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            try
            {
                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType == null)
                {
                    throw new InvalidOperationException(
                        "Microsoft Outlook wurde nicht gefunden.");
                }

                outlookApplication =
                    Activator.CreateInstance(outlookType);

                if (outlookApplication == null)
                {
                    throw new InvalidOperationException(
                        "Outlook konnte nicht gestartet werden.");
                }

                dynamic outlook =
                    outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace("MAPI");

                dynamic outlookNs =
                    outlookNamespace!;

                stores =
                    outlookNs.Stores;

                foreach (SearchResult result in results)
                {
                    object? mailItem = null;

                    try
                    {
                        mailItem =
                            FindOutlookItem(
                                outlookNs,
                                stores,
                                result);

                        if (mailItem == null)
                        {
                            failed++;
                            continue;
                        }

                        dynamic mail =
                            mailItem;

                        string fileName =
                            BuildExportFileName(result);

                        string path =
                            GetUniqueFilePath(
                                targetFolder,
                                fileName);

                        // Outlook OlSaveAsType.olMSG = 3
                        mail.SaveAs(path, 3);

                        exported++;
                    }
                    catch
                    {
                        failed++;
                    }
                    finally
                    {
                        ReleaseComObject(mailItem);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Der Export konnte nicht durchgeführt werden.\n\n" +
                    ex.Message,
                    "Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return;
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }

            string message =
                $"{exported:N0} E-Mail(s) wurden als Outlook-.msg exportiert.";

            if (failed > 0)
            {
                message +=
                    $"\n\n{failed:N0} E-Mail(s) konnten nicht exportiert werden.";
            }

            message +=
                $"\n\nZielordner:\n{targetFolder}";

            MessageBox.Show(
                message,
                "Export abgeschlossen",
                MessageBoxButton.OK,
                failed == 0
                    ? MessageBoxImage.Information
                    : MessageBoxImage.Warning);
        }

        private static string BuildExportFileName(
            SearchResult result)
        {
            string datePart =
                result.SortDate > DateTime.MinValue
                    ? result.SortDate.ToString(
                        "yyyy-MM-dd_HHmm")
                    : "Ohne_Datum";

            string subject =
                string.IsNullOrWhiteSpace(result.Subject)
                    ? "Ohne_Betreff"
                    : result.Subject;

            subject =
                MakeSafeFileName(subject);

            if (subject.Length > 100)
            {
                subject =
                    subject.Substring(0, 100);
            }

            return
                $"{datePart}_{subject}.msg";
        }

        private static string MakeSafeFileName(
            string value)
        {
            foreach (
                char invalidChar
                in Path.GetInvalidFileNameChars())
            {
                value =
                    value.Replace(
                        invalidChar,
                        '_');
            }

            value =
                value.Trim();

            while (value.Contains("__"))
            {
                value =
                    value.Replace(
                        "__",
                        "_");
            }

            return value;
        }

        private static string GetUniqueFilePath(
            string folder,
            string fileName)
        {
            string path =
                Path.Combine(
                    folder,
                    fileName);

            if (!File.Exists(path))
            {
                return path;
            }

            string name =
                Path.GetFileNameWithoutExtension(
                    fileName);

            string extension =
                Path.GetExtension(
                    fileName);

            int counter = 2;

            while (true)
            {
                string candidate =
                    Path.Combine(
                        folder,
                        $"{name}_{counter}{extension}");

                if (!File.Exists(candidate))
                {
                    return candidate;
                }

                counter++;
            }
        }

        // =========================================================
        // BUILD 1130 – INTELLIGENTE SACHVERHALTSERKENNUNG
        // =========================================================

        private void ShowConversationButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            ShowSelectedConversations();
        }

        private void ShowConversationMenuItem_Click(
            object sender,
            RoutedEventArgs e)
        {
            ShowSelectedConversations();
        }

        private void ShowSelectedConversations()
        {
            List<SearchResult> selected =
                GetSelectedSearchResults();

            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Bitte zuerst mindestens eine E-Mail auswählen.",
                    "Sachverhalt erkennen",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            List<IndexRecord> allRecords =
                ReadIndexRecords();

            List<ConversationGroup> allGroups =
                BuildConversationGroups(
                    allRecords);

            HashSet<string> selectedEntryIds =
                selected
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(
                            x.EntryId))
                    .Select(x =>
                        x.EntryId)
                    .ToHashSet(
                        StringComparer.OrdinalIgnoreCase);

            HashSet<string> selectedConversationIds =
                selected
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(
                            x.ConversationId))
                    .Select(x =>
                        x.ConversationId)
                    .ToHashSet(
                        StringComparer.OrdinalIgnoreCase);

            List<ConversationGroup> anchorGroups =
                allGroups
                    .Where(group =>
                        group.Records.Any(record =>
                            selectedEntryIds.Contains(
                                record.EntryId)) ||
                        (!string.IsNullOrWhiteSpace(
                                group.ConversationId) &&
                         selectedConversationIds.Contains(
                             group.ConversationId)))
                    .ToList();

            if (anchorGroups.Count == 0)
            {
                /*
                 * Fallback für sehr alte/ungewöhnliche Einträge.
                 */
                List<SearchResult> fallback =
                    selected
                        .OrderBy(x =>
                            x.SortDate)
                        .ToList();

                foreach (SearchResult mail in fallback)
                {
                    mail.RelationReason =
                        "Manuell ausgewählt";
                }

                ShowConversationWindow(
                    fallback,
                    0,
                    0);

                return;
            }

            HashSet<string> anchorKeys =
                anchorGroups
                    .Select(x => x.GroupKey)
                    .ToHashSet(
                        StringComparer.OrdinalIgnoreCase);

            List<ConversationGroup> acceptedGroups =
                new List<ConversationGroup>();

            /*
             * Die sicher ausgewählten Mailketten sind immer dabei.
             */
            foreach (ConversationGroup anchor in anchorGroups)
            {
                anchor.RelationScore = 100;
                anchor.RelationReason =
                    "Ausgewählte Mailkette";

                acceptedGroups.Add(anchor);
            }

            /*
             * Alle übrigen Mailketten werden gegen die
             * ausgewählten Ausgangsketten bewertet.
             *
             * Schwelle bewusst hoch: 65 Punkte.
             */
            foreach (
                ConversationGroup candidate
                in allGroups)
            {
                if (anchorKeys.Contains(
                    candidate.GroupKey))
                {
                    continue;
                }

                SachverhaltMatch? bestMatch =
                    null;

                foreach (
                    ConversationGroup anchor
                    in anchorGroups)
                {
                    SachverhaltMatch match =
                        CalculateSachverhaltMatch(
                            anchor,
                            candidate);

                    if (bestMatch == null ||
                        match.Score >
                        bestMatch.Score)
                    {
                        bestMatch =
                            match;
                    }
                }

                if (bestMatch != null &&
                    bestMatch.Score >= 65)
                {
                    candidate.RelationScore =
                        bestMatch.Score;

                    candidate.RelationReason =
                        bestMatch.Reason;

                    acceptedGroups.Add(
                        candidate);
                }
            }

            List<SearchResult> mails =
                new List<SearchResult>();

            foreach (
                ConversationGroup group
                in acceptedGroups)
            {
                foreach (
                    IndexRecord record
                    in group.Records)
                {
                    SearchResult result =
                        ConvertRecordToSearchResult(
                            record,
                            "");

                    result.RelationScore =
                        group.RelationScore;

                    result.RelationReason =
                        group.RelationReason;

                    mails.Add(result);
                }
            }

            mails =
                mails
                    .GroupBy(
                        x =>
                            string.IsNullOrWhiteSpace(
                                x.EntryId)
                                ? Guid.NewGuid()
                                    .ToString()
                                : x.EntryId,
                        StringComparer.OrdinalIgnoreCase)
                    .Select(x =>
                        x.First())
                    .OrderBy(x =>
                        x.SortDate)
                    .ToList();

            int additionalGroups =
                acceptedGroups.Count -
                anchorGroups.Count;

            int additionalMails =
                acceptedGroups
                    .Where(x =>
                        !anchorKeys.Contains(
                            x.GroupKey))
                    .Sum(x =>
                        x.Records.Count);

            ShowConversationWindow(
                mails,
                additionalGroups,
                additionalMails);
        }

        private static List<ConversationGroup>
            BuildConversationGroups(
                List<IndexRecord> records)
        {
            List<ConversationGroup> groups =
                new List<ConversationGroup>();

            int singleton = 0;

            foreach (
                IGrouping<string, IndexRecord> grouping
                in records.GroupBy(
                    record =>
                    {
                        if (!string.IsNullOrWhiteSpace(
                            record.ConversationId))
                        {
                            return
                                "CID:" +
                                record.ConversationId;
                        }

                        /*
                         * Mails ohne ConversationID nicht
                         * künstlich miteinander verbinden.
                         */
                        singleton++;

                        return
                            "SINGLE:" +
                            singleton;
                    },
                    StringComparer.OrdinalIgnoreCase))
            {
                List<IndexRecord> groupRecords =
                    grouping.ToList();

                ConversationGroup group =
                    new ConversationGroup
                    {
                        GroupKey =
                            grouping.Key,

                        ConversationId =
                            groupRecords
                                .Select(x =>
                                    x.ConversationId)
                                .FirstOrDefault(x =>
                                    !string.IsNullOrWhiteSpace(x))
                            ??
                            "",

                        Records =
                            groupRecords
                    };

                group.FirstDate =
                    groupRecords
                        .Select(x =>
                            ParseIndexDate(
                                x.Date))
                        .Where(x =>
                            x.HasValue)
                        .Select(x =>
                            x!.Value)
                        .DefaultIfEmpty(
                            DateTime.MinValue)
                        .Min();

                group.LastDate =
                    groupRecords
                        .Select(x =>
                            ParseIndexDate(
                                x.Date))
                        .Where(x =>
                            x.HasValue)
                        .Select(x =>
                            x!.Value)
                        .DefaultIfEmpty(
                            DateTime.MinValue)
                        .Max();

                foreach (
                    IndexRecord record
                    in groupRecords)
                {
                    string normalizedSubject =
                        NormalizeSubject(
                            record.Subject);

                    if (!string.IsNullOrWhiteSpace(
                        normalizedSubject))
                    {
                        group.NormalizedSubjects.Add(
                            normalizedSubject);

                        foreach (
                            string word
                            in GetMeaningfulWords(
                                normalizedSubject))
                        {
                            group.SubjectWords.Add(
                                word);
                        }
                    }

                    foreach (
                        string participant
                        in ExtractParticipants(
                            record.Sender,
                            record.Recipient,
                            record.Cc))
                    {
                        group.Participants.Add(
                            participant);
                    }

                    foreach (
                        string word
                        in GetMeaningfulWords(
                            record.AttachmentNames))
                    {
                        group.AttachmentWords.Add(
                            word);
                    }
                }

                groups.Add(group);
            }

            return groups;
        }

        private static SachverhaltMatch
            CalculateSachverhaltMatch(
                ConversationGroup anchor,
                ConversationGroup candidate)
        {
            int score = 0;

            List<string> reasons =
                new List<string>();

            // -----------------------------------------------------
            // 1. Betreff
            // -----------------------------------------------------

            bool exactSubject =
                anchor.NormalizedSubjects
                    .Any(subject =>
                        candidate.NormalizedSubjects
                            .Contains(subject));

            double subjectSimilarity =
                CalculateJaccard(
                    anchor.SubjectWords,
                    candidate.SubjectWords);

            if (exactSubject &&
                anchor.NormalizedSubjects.Count > 0 &&
                candidate.NormalizedSubjects.Count > 0)
            {
                score += 55;

                reasons.Add(
                    "gleicher bereinigter Betreff");
            }
            else if (subjectSimilarity >= 0.75)
            {
                score += 45;

                reasons.Add(
                    "sehr ähnlicher Betreff");
            }
            else if (subjectSimilarity >= 0.50)
            {
                score += 30;

                reasons.Add(
                    "ähnlicher Betreff");
            }
            else if (subjectSimilarity >= 0.35)
            {
                score += 15;

                reasons.Add(
                    "teilweise ähnlicher Betreff");
            }

            // -----------------------------------------------------
            // 2. Beteiligte
            // -----------------------------------------------------

            double participantSimilarity =
                CalculateJaccard(
                    anchor.Participants,
                    candidate.Participants);

            int commonParticipants =
                anchor.Participants
                    .Intersect(
                        candidate.Participants,
                        StringComparer.OrdinalIgnoreCase)
                    .Count();

            if (participantSimilarity >= 0.60 &&
                commonParticipants >= 2)
            {
                score += 25;

                reasons.Add(
                    "stark überschneidende Beteiligte");
            }
            else if (commonParticipants >= 2)
            {
                score += 18;

                reasons.Add(
                    "mehrere gemeinsame Beteiligte");
            }
            else if (commonParticipants == 1)
            {
                score += 8;

                reasons.Add(
                    "gemeinsamer Beteiligter");
            }

            // -----------------------------------------------------
            // 3. Zeitraum
            // -----------------------------------------------------

            double dayDistance =
                CalculateGroupDayDistance(
                    anchor,
                    candidate);

            if (dayDistance <= 7)
            {
                score += 18;

                reasons.Add(
                    "sehr zeitnah");
            }
            else if (dayDistance <= 30)
            {
                score += 12;

                reasons.Add(
                    "zeitliche Nähe");
            }
            else if (dayDistance <= 90)
            {
                score += 6;

                reasons.Add(
                    "zeitlicher Zusammenhang");
            }

            // -----------------------------------------------------
            // 4. Gemeinsame Dokument-/Anhangbegriffe
            // -----------------------------------------------------

            double attachmentSimilarity =
                CalculateJaccard(
                    anchor.AttachmentWords,
                    candidate.AttachmentWords);

            if (attachmentSimilarity >= 0.60 &&
                anchor.AttachmentWords.Count > 0 &&
                candidate.AttachmentWords.Count > 0)
            {
                score += 12;

                reasons.Add(
                    "ähnliche Anhänge");
            }
            else if (attachmentSimilarity >= 0.35)
            {
                score += 6;

                reasons.Add(
                    "gemeinsame Anhangbegriffe");
            }

            /*
             * Sicherheitsregel:
             * Nur gemeinsame Beteiligte + zeitliche Nähe reichen
             * NICHT für einen Sachverhalt.
             *
             * Ohne nennenswerte Betreffähnlichkeit wird der
             * Maximalwert begrenzt.
             */
            if (subjectSimilarity < 0.35 &&
                !exactSubject)
            {
                score =
                    Math.Min(
                        score,
                        50);
            }

            return new SachverhaltMatch
            {
                Score =
                    score,

                Reason =
                    reasons.Count > 0
                        ? string.Join(
                            " + ",
                            reasons)
                        : "kein ausreichender Zusammenhang"
            };
        }

        private static string NormalizeSubject(
            string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return "";
            }

            string result =
                subject.Trim();

            bool changed;

            do
            {
                changed = false;

                string[] prefixes =
                {
                    "RE:",
                    "AW:",
                    "WG:",
                    "FW:",
                    "FWD:",
                    "ANTWORT:",
                    "WEITERLEITUNG:"
                };

                foreach (string prefix in prefixes)
                {
                    if (result.StartsWith(
                        prefix,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        result =
                            result.Substring(
                                prefix.Length)
                            .Trim();

                        changed = true;
                    }
                }
            }
            while (changed);

            result =
                result
                    .Replace(
                        "_",
                        " ")
                    .Replace(
                        "-",
                        " ")
                    .Replace(
                        "/",
                        " ")
                    .Replace(
                        "\\",
                        " ");

            while (result.Contains("  "))
            {
                result =
                    result.Replace(
                        "  ",
                        " ");
            }

            return
                result
                    .Trim()
                    .ToLowerInvariant();
        }

        private static HashSet<string>
            GetMeaningfulWords(
                string text)
        {
            HashSet<string> words =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(text))
            {
                return words;
            }

            char[] separators =
            {
                ' ',
                '\t',
                '\r',
                '\n',
                '.',
                ',',
                ';',
                ':',
                '-',
                '_',
                '/',
                '\\',
                '(',
                ')',
                '[',
                ']',
                '{',
                '}',
                '!',
                '?',
                '"',
                '\'',
                '+',
                '=',
                '&'
            };

            HashSet<string> stopWords =
                new HashSet<string>(
                    new[]
                    {
                        "der", "die", "das",
                        "den", "dem", "des",
                        "ein", "eine", "einer",
                        "eines", "einem", "einen",
                        "und", "oder", "mit",
                        "für", "von", "vom",
                        "im", "in", "am", "an",
                        "auf", "zu", "zum", "zur",
                        "bei", "aus", "über",
                        "unter", "wegen", "nach",
                        "vor", "bitte", "mail",
                        "email", "re", "aw", "wg",
                        "fw", "fwd", "pdf", "docx",
                        "xlsx", "txt", "csv"
                    },
                    StringComparer.OrdinalIgnoreCase);

            foreach (
                string raw
                in text.Split(
                    separators,
                    StringSplitOptions.RemoveEmptyEntries |
                    StringSplitOptions.TrimEntries))
            {
                string word =
                    raw.Trim()
                       .ToLowerInvariant();

                if (word.Length < 3)
                {
                    continue;
                }

                if (stopWords.Contains(word))
                {
                    continue;
                }

                words.Add(word);
            }

            return words;
        }

        private static HashSet<string>
            ExtractParticipants(
                params string[] values)
        {
            HashSet<string> result =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            char[] separators =
            {
                ';',
                ',',
                '\r',
                '\n'
            };

            foreach (string value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                foreach (
                    string part
                    in value.Split(
                        separators,
                        StringSplitOptions.RemoveEmptyEntries |
                        StringSplitOptions.TrimEntries))
                {
                    string participant =
                        part.Trim()
                            .ToLowerInvariant();

                    if (participant.Length < 2)
                    {
                        continue;
                    }

                    result.Add(participant);
                }
            }

            return result;
        }

        private static double CalculateJaccard(
            HashSet<string> first,
            HashSet<string> second)
        {
            if (first.Count == 0 ||
                second.Count == 0)
            {
                return 0;
            }

            int intersection =
                first
                    .Intersect(
                        second,
                        StringComparer.OrdinalIgnoreCase)
                    .Count();

            int union =
                first
                    .Union(
                        second,
                        StringComparer.OrdinalIgnoreCase)
                    .Count();

            if (union == 0)
            {
                return 0;
            }

            return
                (double)intersection /
                union;
        }

        private static double CalculateGroupDayDistance(
            ConversationGroup first,
            ConversationGroup second)
        {
            if (first.FirstDate == DateTime.MinValue ||
                second.FirstDate == DateTime.MinValue)
            {
                return double.MaxValue;
            }

            /*
             * Zeiträume überschneiden sich.
             */
            if (first.FirstDate <= second.LastDate &&
                second.FirstDate <= first.LastDate)
            {
                return 0;
            }

            if (first.LastDate < second.FirstDate)
            {
                return
                    (second.FirstDate -
                     first.LastDate)
                    .TotalDays;
            }

            return
                (first.FirstDate -
                 second.LastDate)
                .TotalDays;
        }

        // =========================================================
        // SACHVERHALTSFENSTER
        // =========================================================

        private void ShowConversationWindow(
            List<SearchResult> mails,
            int additionalGroups,
            int additionalMails)
        {
            if (mails.Count == 0)
            {
                MessageBox.Show(
                    "Für die ausgewählte Mail wurde kein Sachverhalt gefunden.",
                    "Sachverhalt",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            int conversationCount =
                mails
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(
                            x.ConversationId))
                    .Select(x =>
                        x.ConversationId)
                    .Distinct(
                        StringComparer.OrdinalIgnoreCase)
                    .Count();

            DateTime firstDate =
                mails
                    .Where(x =>
                        x.SortDate > DateTime.MinValue)
                    .Select(x =>
                        x.SortDate)
                    .DefaultIfEmpty(
                        DateTime.MinValue)
                    .Min();

            DateTime lastDate =
                mails
                    .Where(x =>
                        x.SortDate > DateTime.MinValue)
                    .Select(x =>
                        x.SortDate)
                    .DefaultIfEmpty(
                        DateTime.MinValue)
                    .Max();

            Window window =
                new Window
                {
                    Title =
                        "Datenfinder – Sachverhalt",

                    Width =
                        1350,

                    Height =
                        720,

                    MinWidth =
                        950,

                    MinHeight =
                        520,

                    WindowStartupLocation =
                        WindowStartupLocation.CenterOwner,

                    Owner =
                        this,

                    Background =
                        new SolidColorBrush(
                            Color.FromRgb(
                                244,
                                246,
                                248))
                };

            Grid root =
                new Grid
                {
                    Margin =
                        new Thickness(18)
                };

            root.RowDefinitions.Add(
                new RowDefinition
                {
                    Height =
                        GridLength.Auto
                });

            root.RowDefinitions.Add(
                new RowDefinition
                {
                    Height =
                        new GridLength(10)
                });

            root.RowDefinitions.Add(
                new RowDefinition
                {
                    Height =
                        new GridLength(
                            1,
                            GridUnitType.Star)
                });

            root.RowDefinitions.Add(
                new RowDefinition
                {
                    Height =
                        new GridLength(10)
                });

            root.RowDefinitions.Add(
                new RowDefinition
                {
                    Height =
                        GridLength.Auto
                });

            Border summaryBorder =
                new Border
                {
                    Background =
                        Brushes.White,

                    BorderBrush =
                        new SolidColorBrush(
                            Color.FromRgb(
                                210,
                                216,
                                222)),

                    BorderThickness =
                        new Thickness(1),

                    CornerRadius =
                        new CornerRadius(7),

                    Padding =
                        new Thickness(14)
                };

            StackPanel summary =
                new StackPanel();

            summary.Children.Add(
                new TextBlock
                {
                    Text =
                        "Erkannter Sachverhalt",

                    FontSize =
                        20,

                    FontWeight =
                        FontWeights.SemiBold,

                    Foreground =
                        new SolidColorBrush(
                            Color.FromRgb(
                                18,
                                59,
                                99))
                });

            string dateText =
                firstDate > DateTime.MinValue
                    ? $"{firstDate:dd.MM.yyyy} bis {lastDate:dd.MM.yyyy}"
                    : "Zeitraum nicht verfügbar";

            summary.Children.Add(
                new TextBlock
                {
                    Text =
                        $"{mails.Count:N0} E-Mails  •  " +
                        $"{Math.Max(1, conversationCount):N0} Mailketten  •  " +
                        $"{dateText}",

                    Margin =
                        new Thickness(
                            0,
                            5,
                            0,
                            0),

                    Foreground =
                        Brushes.DimGray,

                    FontSize =
                        12
                });

            if (additionalGroups > 0)
            {
                summary.Children.Add(
                    new TextBlock
                    {
                        Text =
                            $"Zusätzlich erkannt: {additionalGroups:N0} weitere Mailkette(n) mit {additionalMails:N0} E-Mail(s).",

                        Margin =
                            new Thickness(
                                0,
                                5,
                                0,
                                0),

                        Foreground =
                            new SolidColorBrush(
                                Color.FromRgb(
                                    0,
                                    120,
                                    70)),

                        FontWeight =
                            FontWeights.SemiBold,

                        FontSize =
                            11
                    });
            }
            else
            {
                summary.Children.Add(
                    new TextBlock
                    {
                        Text =
                            "Es wurden keine zusätzlichen ausreichend sicheren Mailketten erkannt.",

                        Margin =
                            new Thickness(
                                0,
                                5,
                                0,
                                0),

                        Foreground =
                            Brushes.Gray,

                        FontSize =
                            10
                    });
            }

            summary.Children.Add(
                new TextBlock
                {
                    Text =
                        "Automatische Zuordnungen werden konservativ anhand von Betreff, Beteiligten, Zeitraum und Anhängen bewertet.",

                    Margin =
                        new Thickness(
                            0,
                            4,
                            0,
                            0),

                    Foreground =
                        Brushes.Gray,

                    FontSize =
                        10,

                    TextWrapping =
                        TextWrapping.Wrap
                });

            summaryBorder.Child =
                summary;

            Grid.SetRow(
                summaryBorder,
                0);

            root.Children.Add(
                summaryBorder);

            DataGrid conversationGrid =
                new DataGrid
                {
                    AutoGenerateColumns =
                        false,

                    IsReadOnly =
                        true,

                    CanUserAddRows =
                        false,

                    CanUserDeleteRows =
                        false,

                    SelectionMode =
                        DataGridSelectionMode.Extended,

                    SelectionUnit =
                        DataGridSelectionUnit.FullRow,

                    GridLinesVisibility =
                        DataGridGridLinesVisibility.Horizontal,

                    Background =
                        Brushes.White,

                    ItemsSource =
                        mails,

                    HorizontalScrollBarVisibility =
                        ScrollBarVisibility.Auto,

                    VerticalScrollBarVisibility =
                        ScrollBarVisibility.Auto
                };

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Datum",
                    Binding =
                        new System.Windows.Data.Binding(
                            "Date"),
                    Width = 125
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Absender",
                    Binding =
                        new System.Windows.Data.Binding(
                            "Sender"),
                    Width = 175
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Empfänger",
                    Binding =
                        new System.Windows.Data.Binding(
                            "Recipient"),
                    Width = 190
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Betreff",
                    Binding =
                        new System.Windows.Data.Binding(
                            "Subject"),
                    Width =
                        new DataGridLength(
                            2,
                            DataGridLengthUnitType.Star)
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Zusammenhang",
                    Binding =
                        new System.Windows.Data.Binding(
                            "RelationReason"),
                    Width = 260
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Bewertung",
                    Binding =
                        new System.Windows.Data.Binding(
                            "RelationScore"),
                    Width = 75
                });

            conversationGrid.Columns.Add(
                new DataGridTextColumn
                {
                    Header = "Anhänge",
                    Binding =
                        new System.Windows.Data.Binding(
                            "AttachmentNames"),
                    Width = 190
                });

            conversationGrid.MouseDoubleClick +=
                (_, _) =>
                {
                    if (conversationGrid.SelectedItem
                        is SearchResult result)
                    {
                        OpenSpecificMail(result);
                    }
                };

            Grid.SetRow(
                conversationGrid,
                2);

            root.Children.Add(
                conversationGrid);

            Grid buttonGrid =
                new Grid();

            buttonGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width =
                        new GridLength(
                            1,
                            GridUnitType.Star)
                });

            buttonGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width =
                        GridLength.Auto
                });

            buttonGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width =
                        GridLength.Auto
                });

            buttonGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width =
                        GridLength.Auto
                });

            buttonGrid.Children.Add(
                new TextBlock
                {
                    Text =
                        "Doppelklick öffnet die Original-Mail in Outlook.",

                    VerticalAlignment =
                        VerticalAlignment.Center,

                    Foreground =
                        Brushes.Gray,

                    FontSize =
                        11
                });

            Button openButton =
                new Button
                {
                    Content =
                        "Original-Mail öffnen",

                    Width =
                        145,

                    Height =
                        32,

                    Margin =
                        new Thickness(
                            5,
                            0,
                            0,
                            0)
                };

            Grid.SetColumn(
                openButton,
                1);

            openButton.Click +=
                (_, _) =>
                {
                    if (conversationGrid.SelectedItem
                        is not SearchResult result)
                    {
                        MessageBox.Show(
                            window,
                            "Bitte zuerst eine E-Mail auswählen.",
                            "Original-Mail öffnen",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);

                        return;
                    }

                    OpenSpecificMail(result);
                };

            buttonGrid.Children.Add(
                openButton);

            Button exportSelectionButton =
                new Button
                {
                    Content =
                        "Auswahl exportieren",

                    Width =
                        140,

                    Height =
                        32,

                    Margin =
                        new Thickness(
                            5,
                            0,
                            0,
                            0)
                };

            Grid.SetColumn(
                exportSelectionButton,
                2);

            exportSelectionButton.Click +=
                (_, _) =>
                {
                    List<SearchResult> selected =
                        conversationGrid
                            .SelectedItems
                            .OfType<SearchResult>()
                            .ToList();

                    if (selected.Count == 0)
                    {
                        MessageBox.Show(
                            window,
                            "Bitte zuerst eine oder mehrere E-Mails auswählen.",
                            "Export",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);

                        return;
                    }

                    ExportSearchResults(selected);
                };

            buttonGrid.Children.Add(
                exportSelectionButton);

            Button exportAllButton =
                new Button
                {
                    Content =
                        "Gesamten Sachverhalt exportieren",

                    Width =
                        205,

                    Height =
                        32,

                    Margin =
                        new Thickness(
                            5,
                            0,
                            0,
                            0),

                    FontWeight =
                        FontWeights.SemiBold,

                    Foreground =
                        Brushes.White,

                    Background =
                        new SolidColorBrush(
                            Color.FromRgb(
                                82,
                                109,
                                130))
                };

            Grid.SetColumn(
                exportAllButton,
                3);

            exportAllButton.Click +=
                (_, _) =>
                {
                    ExportSearchResults(mails);
                };

            buttonGrid.Children.Add(
                exportAllButton);

            Grid.SetRow(
                buttonGrid,
                4);

            root.Children.Add(
                buttonGrid);

            window.Content =
                root;

            window.ShowDialog();
        }

        // =========================================================
        // ORIGINAL-MAIL ÖFFNEN
        // =========================================================

        private void OpenSelectedMailButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            List<SearchResult> selected =
                GetSelectedSearchResults();

            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Bitte zuerst eine E-Mail auswählen.",
                    "Original-Mail öffnen",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            OpenSpecificMail(
                selected[0]);
        }

        private void OpenOriginalMailMenuItem_Click(
            object sender,
            RoutedEventArgs e)
        {
            OpenSelectedMailButton_Click(
                sender,
                e);
        }

        private void SearchResultsGrid_MouseDoubleClick(
            object sender,
            MouseButtonEventArgs e)
        {
            if (SearchResultsGrid.SelectedItem
                is SearchResult result)
            {
                OpenSpecificMail(result);
            }
        }

        private void OpenSpecificMail(
            SearchResult result)
        {
            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;
            object? mailItem = null;

            try
            {
                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType == null)
                {
                    throw new InvalidOperationException(
                        "Das klassische Microsoft Outlook wurde auf diesem PC nicht gefunden.");
                }

                outlookApplication =
                    Activator.CreateInstance(
                        outlookType);

                if (outlookApplication == null)
                {
                    throw new InvalidOperationException(
                        "Outlook konnte nicht gestartet werden.");
                }

                dynamic outlook =
                    outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace(
                        "MAPI");

                dynamic outlookNs =
                    outlookNamespace!;

                stores =
                    outlookNs.Stores;

                mailItem =
                    FindOutlookItem(
                        outlookNs,
                        stores,
                        result);

                if (mailItem == null)
                {
                    MessageBox.Show(
                        "Die Original-Mail konnte in Outlook nicht mehr gefunden werden.\n\n" +
                        "Möglicherweise wurde sie seit der letzten Index-Aktualisierung verschoben oder gelöscht.",
                        "Original-Mail nicht gefunden",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    return;
                }

                dynamic mail =
                    mailItem;

                mail.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Die Original-Mail konnte nicht geöffnet werden.\n\n" +
                    ex.Message,
                    "Outlook",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                ReleaseComObject(mailItem);
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private static object? FindOutlookItem(
            dynamic outlookNamespace,
            object? storesObject,
            SearchResult result)
        {
            object? mailItem = null;

            if (storesObject != null &&
                !string.IsNullOrWhiteSpace(
                    result.Mailbox))
            {
                dynamic stores =
                    storesObject;

                int storeCount =
                    stores.Count;

                for (int i = 1;
                     i <= storeCount;
                     i++)
                {
                    object? storeObject = null;

                    try
                    {
                        storeObject =
                            stores.Item(i);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (!storeName.Equals(
                            result.Mailbox,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string storeId = "";

                        try
                        {
                            storeId =
                                store.StoreID ??
                                "";
                        }
                        catch
                        {
                        }

                        if (!string.IsNullOrWhiteSpace(
                            storeId))
                        {
                            try
                            {
                                mailItem =
                                    outlookNamespace.GetItemFromID(
                                        result.EntryId,
                                        storeId);
                            }
                            catch
                            {
                                mailItem = null;
                            }
                        }

                        break;
                    }
                    finally
                    {
                        ReleaseComObject(
                            storeObject);
                    }
                }
            }

            if (mailItem == null)
            {
                try
                {
                    mailItem =
                        outlookNamespace.GetItemFromID(
                            result.EntryId);
                }
                catch
                {
                    mailItem = null;
                }
            }

            return mailItem;
        }

        // =========================================================
        // OUTLOOK-HILFSMETHODEN
        // =========================================================

        private static string GetFlagDescription(
            dynamic item)
        {
            try
            {
                int flagStatus =
                    (int)item.FlagStatus;

                string flagRequest = "";

                try
                {
                    flagRequest =
                        item.FlagRequest ??
                        "";
                }
                catch
                {
                }

                if (flagStatus == 1)
                {
                    return
                        string.IsNullOrWhiteSpace(
                            flagRequest)
                            ? "Erledigt / Häkchen"
                            : "Erledigt / Häkchen – " +
                              flagRequest;
                }

                if (flagStatus == 2)
                {
                    return
                        string.IsNullOrWhiteSpace(
                            flagRequest)
                            ? "Fähnchen"
                            : "Fähnchen – " +
                              flagRequest;
                }

                return
                    "Keine Kennzeichnung";
            }
            catch
            {
                return
                    "Keine Kennzeichnung";
            }
        }

        private static string SafeDynamicString(
            dynamic item,
            string propertyName)
        {
            try
            {
                return propertyName switch
                {
                    "Subject" =>
                        item.Subject ?? "",

                    "SenderName" =>
                        item.SenderName ?? "",

                    "Body" =>
                        item.Body ?? "",

                    "DisplayName" =>
                        item.DisplayName ?? "",

                    "Name" =>
                        item.Name ?? "",

                    "To" =>
                        item.To ?? "",

                    "CC" =>
                        item.CC ?? "",

                    "Categories" =>
                        item.Categories ?? "",

                    "ConversationID" =>
                        item.ConversationID ?? "",

                    "EntryID" =>
                        item.EntryID ?? "",

                    _ =>
                        ""
                };
            }
            catch
            {
                return "";
            }
        }

        private static string SafeDynamicDateTime(
            dynamic item)
        {
            try
            {
                DateTime value =
                    item.ReceivedTime;

                if (value.Year > 1900)
                {
                    return value.ToString(
                        "yyyy-MM-dd HH:mm:ss",
                        CultureInfo.InvariantCulture);
                }
            }
            catch
            {
            }

            try
            {
                DateTime value =
                    item.SentOn;

                return value.ToString(
                    "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture);
            }
            catch
            {
                return "";
            }
        }

        private static DateTime?
            ParseIndexDate(
                string value)
        {
            if (DateTime.TryParseExact(
                value,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime result))
            {
                return result;
            }

            return null;
        }

        private static string GetComboBoxText(
            ComboBox comboBox)
        {
            if (comboBox.SelectedItem
                is ComboBoxItem item)
            {
                return
                    item.Content
                        ?.ToString()
                    ??
                    "";
            }

            return
                comboBox.SelectedItem
                    ?.ToString()
                ??
                "";
        }

        private static string CleanIndexText(
            string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return text
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();
        }

        private static string FormatFileSize(
            long bytes)
        {
            if (bytes >=
                1024L *
                1024L *
                1024L)
            {
                return
                    $"{bytes / (1024.0 * 1024.0 * 1024.0):0.0} GB";
            }

            if (bytes >=
                1024L *
                1024L)
            {
                return
                    $"{bytes / (1024.0 * 1024.0):0.0} MB";
            }

            if (bytes >= 1024L)
            {
                return
                    $"{bytes / 1024.0:0.0} KB";
            }

            return
                $"{bytes} Byte";
        }

        private static async Task RefreshUi()
        {
            await Application.Current
                .Dispatcher
                .InvokeAsync(
                    () =>
                    {
                    },
                    DispatcherPriority.Background);
        }

        private static void ReleaseComObject(
            object? comObject)
        {
            if (comObject != null &&
                Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(
                    comObject);
            }
        }

        // =========================================================
        // DATENKLASSEN
        // =========================================================

        private class ConversationGroup
        {
            public string GroupKey { get; set; } = "";
            public string ConversationId { get; set; } = "";

            public List<IndexRecord> Records { get; set; } =
                new List<IndexRecord>();

            public HashSet<string> NormalizedSubjects { get; set; } =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            public HashSet<string> SubjectWords { get; set; } =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            public HashSet<string> Participants { get; set; } =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            public HashSet<string> AttachmentWords { get; set; } =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            public DateTime FirstDate { get; set; }
            public DateTime LastDate { get; set; }

            public int RelationScore { get; set; }
            public string RelationReason { get; set; } = "";
        }

        private class SachverhaltMatch
        {
            public int Score { get; set; }
            public string Reason { get; set; } = "";
        }

        private class AttachmentIndexData
        {
            public bool HasAttachments { get; set; }

            public string Names { get; set; } = "";

            public string SearchText { get; set; } = "";
        }

        private class IndexRecord
        {
            public string Date { get; set; } = "";
            public string Sender { get; set; } = "";
            public string Recipient { get; set; } = "";
            public string Cc { get; set; } = "";
            public string Mailbox { get; set; } = "";
            public string Subject { get; set; } = "";
            public string Folder { get; set; } = "";
            public string Flag { get; set; } = "";
            public string Attachment { get; set; } = "";
            public string Categories { get; set; } = "";
            public string ConversationId { get; set; } = "";
            public string EntryId { get; set; } = "";
            public string Body { get; set; } = "";
            public string AttachmentNames { get; set; } = "";
            public string AttachmentText { get; set; } = "";

            public bool ContentEquals(
                IndexRecord other)
            {
                return
                    Date == other.Date &&
                    Sender == other.Sender &&
                    Recipient == other.Recipient &&
                    Cc == other.Cc &&
                    Mailbox == other.Mailbox &&
                    Subject == other.Subject &&
                    Folder == other.Folder &&
                    Flag == other.Flag &&
                    Attachment == other.Attachment &&
                    Categories == other.Categories &&
                    ConversationId == other.ConversationId &&
                    EntryId == other.EntryId &&
                    Body == other.Body &&
                    AttachmentNames == other.AttachmentNames &&
                    AttachmentText == other.AttachmentText;
            }
        }

        private class SearchOptions
        {
            public string Query { get; set; } = "";

            public DateTime? FromDate { get; set; }
            public DateTime? ToDate { get; set; }

            public string Mailbox { get; set; } =
                "Alle Postfächer";

            public string Attachment { get; set; } =
                "Alle";

            public string Flag { get; set; } =
                "Alle";

            public string SenderFilter { get; set; } = "";
            public string RecipientFilter { get; set; } = "";

            public bool SearchSubject { get; set; }
            public bool SearchBody { get; set; }
            public bool SearchAttachments { get; set; }

            public string Sort { get; set; } =
                "Neueste zuerst";
        }

        private class SearchResponse
        {
            public List<SearchResult> Results { get; set; } =
                new List<SearchResult>();

            public int TotalMatches { get; set; }

            public bool WasLimited { get; set; }
        }

        public class SearchResult
        {
            public DateTime SortDate { get; set; }

            public string Date { get; set; } = "";
            public string Mailbox { get; set; } = "";
            public string Recipient { get; set; } = "";
            public string Cc { get; set; } = "";
            public string Sender { get; set; } = "";
            public string Flag { get; set; } = "";
            public string Attachment { get; set; } = "";
            public string AttachmentNames { get; set; } = "";
            public string FoundIn { get; set; } = "";
            public string Subject { get; set; } = "";
            public string Folder { get; set; } = "";
            public string Categories { get; set; } = "";
            public string ConversationId { get; set; } = "";
            public string EntryId { get; set; } = "";
            public string Body { get; set; } = "";

            public int RelationScore { get; set; }

            public string RelationReason { get; set; } = "";
        }
    }
}