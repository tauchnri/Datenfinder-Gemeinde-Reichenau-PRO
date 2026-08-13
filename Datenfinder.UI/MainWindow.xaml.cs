using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Datenfinder.UI
{
    public partial class MainWindow : Window
    {
        private const int MailItemClass = 43;
        private const int MaximumSearchResults = 500;
        private const string IndexSchema = "1060";
        private const string PreviousIndexSchema = "1050";

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

        public MainWindow()
        {
            InitializeComponent();

            _indexFolder = Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments),
                "Datenfinder Gemeinde Reichenau PRO");

            _indexPath = Path.Combine(
                _indexFolder,
                "Outlook-Inhaltsindex.txt");

            _syncInfoPath = Path.Combine(
                _indexFolder,
                "Outlook-Sync.txt");

            InitializeFilters();
            CheckExistingIndex();

            _automaticUpdateTimer =
                new DispatcherTimer
                {
                    Interval = TimeSpan.FromHours(1)
                };

            _automaticUpdateTimer.Tick +=
                AutomaticUpdateTimer_Tick;

            _automaticUpdateTimer.Start();
        }

        private void InitializeFilters()
        {
            AttachmentComboBox.SelectedIndex = 0;
            FlagComboBox.SelectedIndex = 0;
            SortComboBox.SelectedIndex = 0;

            MailboxComboBox.Items.Clear();
            MailboxComboBox.Items.Add("Alle Postfächer");
            MailboxComboBox.SelectedIndex = 0;

            SenderFilterTextBox.Text = "";
            RecipientFilterTextBox.Text = "";

            ActiveFiltersText.Text =
                "Aktive Filter: keine";
        }

        private void CheckExistingIndex()
        {
            if (!File.Exists(_indexPath))
            {
                SearchButton.IsEnabled = false;
                CreateIndexButton.Content = "Index erstellen";

                IndexStatusText.Text =
                    "Noch kein Suchindex vorhanden";

                IndexDetailsText.Text =
                    "Outlook prüfen und anschließend den ersten Index erstellen.";

                LastUpdateText.Text = "";

                return;
            }

            if (!IsSearchableIndex())
            {
                SearchButton.IsEnabled = false;

                IndexStatusText.Text =
                    "Vorhandener Suchindex ist nicht kompatibel";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 100, 0));

                IndexDetailsText.Text =
                    "Der Index muss einmal vollständig neu erstellt werden.";

                CreateIndexButton.Content = "Index erstellen";

                return;
            }

            SearchButton.IsEnabled = true;
            CreateIndexButton.Content = "Jetzt aktualisieren";

            IndexStatusText.Text =
                GetIndexSchema() == PreviousIndexSchema
                    ? "Build-1050-Index vorhanden – bereit für intelligente Aktualisierung"
                    : "Outlook-Inhaltsindex ist bereit";

            IndexStatusText.Foreground =
                new SolidColorBrush(
                    Color.FromRgb(0, 120, 70));

            FileInfo fileInfo =
                new FileInfo(_indexPath);

            IndexDetailsText.Text =
                $"Indexgröße: {FormatFileSize(fileInfo.Length)}";

            SearchStatusText.Text =
                "Suchbegriff eingeben oder die Filter verwenden.";

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

                for (int i = 0; i < 15; i++)
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
                            .Substring("Schema:".Length)
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

            return schema == IndexSchema ||
                   schema == PreviousIndexSchema;
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

                while ((line = reader.ReadLine()) != null)
                {
                    string[] columns =
                        line.Split(
                            new[] { '\t' },
                            14,
                            StringSplitOptions.None);

                    if (columns.Length != 14)
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
                        mailboxes.Add(mailbox);
                    }
                }

                string? selected =
                    MailboxComboBox.SelectedItem?.ToString();

                MailboxComboBox.Items.Clear();
                MailboxComboBox.Items.Add(
                    "Alle Postfächer");

                foreach (string mailbox
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

                MailboxComboBox.SelectedIndex = 0;
            }
        }

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
                    "Noch keine inkrementelle Aktualisierung durchgeführt.";
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

        private async void AutomaticUpdateTimer_Tick(
            object? sender,
            EventArgs e)
        {
            if (_updateInProgress)
            {
                return;
            }

            if (!IsSearchableIndex())
            {
                return;
            }

            await UpdateIndexIncrementallyAsync(
                false);
        }

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
                        Color.FromRgb(85, 85, 85));

                OutlookDetailsText.Text = "";
                MailboxNamesText.Text = "";

                CreateIndexButton.IsEnabled = false;

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
                    object? storeObject = null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(i);

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
                        Color.FromRgb(0, 120, 70));

                OutlookDetailsText.Text =
                    $"Gefundene Datenspeicher/Postfächer: {storeCount}";

                MailboxNamesText.Text =
                    storeNames.Count > 0
                        ? "Verbunden: " +
                          string.Join(
                              "  •  ",
                              storeNames)
                        : "";

                CreateIndexButton.IsEnabled = true;

                if (IsSearchableIndex())
                {
                    IndexStatusText.Text =
                        "Index vorhanden – intelligente Aktualisierung bereit";

                    IndexDetailsText.Text =
                        "Nur neue oder geänderte Nachrichten werden aus Outlook neu eingelesen.";

                    CreateIndexButton.Content =
                        "Jetzt aktualisieren";

                    SearchButton.IsEnabled = true;
                }
                else
                {
                    IndexStatusText.Text =
                        "Erster Suchindex muss erstellt werden";

                    IndexDetailsText.Text =
                        "Beim ersten Lauf wird der vollständige Outlook-Bestand indiziert.";

                    CreateIndexButton.Content =
                        "Index erstellen";

                    SearchButton.IsEnabled = false;
                }

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));
            }
            catch (Exception ex)
            {
                OutlookStatusText.Text =
                    "Status: Outlook-Verbindung fehlgeschlagen";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 40, 40));

                OutlookDetailsText.Text =
                    ex.Message;

                CreateIndexButton.IsEnabled =
                    false;
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private async void CreateIndexButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (_updateInProgress)
            {
                return;
            }

            if (IsSearchableIndex())
            {
                await UpdateIndexIncrementallyAsync(
                    true);
            }
            else
            {
                await CreateFullIndexAsync();
            }
        }

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
                CreateIndexButton.IsEnabled = false;
                ConnectOutlookButton.IsEnabled = false;

                ProgressPanel.Visibility =
                    Visibility.Visible;

                IndexProgressBar.IsIndeterminate =
                    true;

                IndexProgressBar.Value = 0;

                ProgressPercentText.Text = "";

                ProgressPhaseText.Text =
                    "Index wird intelligent aktualisiert";

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

                // Sicherheitsüberlappung:
                // Die letzten 10 Minuten werden nochmals geprüft.
                DateTime scanSince =
                    lastSync.AddMinutes(-10);

                _incrementalNewCount = 0;
                _incrementalChangedCount = 0;

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

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject = null;
                    object? rootFolderObject = null;

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

                        await RefreshUi();

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject != null)
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

                ProgressPhaseText.Text =
                    "Indexdatei wird gespeichert";

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

                IndexProgressBar.IsIndeterminate =
                    false;

                IndexProgressBar.Maximum = 100;
                IndexProgressBar.Value = 100;

                ProgressPercentText.Text =
                    "100 %";

                ProgressPhaseText.Text =
                    "Aktualisierung abgeschlossen";

                ProgressCountText.Text =
                    $"{_incrementalNewCount:N0} neue | " +
                    $"{_incrementalChangedCount:N0} geänderte E-Mails";

                ProgressFolderText.Text =
                    $"Index enthält jetzt {records.Count:N0} E-Mails.";

                IndexStatusText.Text =
                    "Outlook-Inhaltsindex ist aktuell";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                IndexDetailsText.Text =
                    $"{storeCount} Postfächer | " +
                    $"{records.Count:N0} E-Mails | " +
                    $"vorher {oldCount:N0}";

                SearchButton.IsEnabled = true;

                CreateIndexButton.Content =
                    "Jetzt aktualisieren";

                LoadMailboxesFromIndex();
                UpdateLastSyncDisplay();

                if (userRequested)
                {
                    SearchStatusText.Text =
                        _incrementalNewCount == 0 &&
                        _incrementalChangedCount == 0
                            ? "Index ist bereits aktuell. Keine neuen oder geänderten E-Mails gefunden."
                            : $"Index aktualisiert: {_incrementalNewCount:N0} neu, " +
                              $"{_incrementalChangedCount:N0} geändert.";
                }

                await Task.Delay(
                    700);

                ProgressPanel.Visibility =
                    Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                IndexProgressBar.IsIndeterminate =
                    false;

                IndexStatusText.Text =
                    "Index-Aktualisierung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 40, 40));

                IndexDetailsText.Text =
                    ex.Message;

                ProgressPhaseText.Text =
                    "Aktualisierung abgebrochen";

                SearchButton.IsEnabled =
                    IsSearchableIndex();
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);

                CreateIndexButton.IsEnabled = true;
                ConnectOutlookButton.IsEnabled = true;

                _updateInProgress = false;
            }
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
                return DateTime.Now.AddDays(-1);
            }
        }

        private async Task UpdateFolderIncrementallyAsync(
            object folderObject,
            string storeName,
            string parentPath,
            DateTime scanSince,
            Dictionary<string, IndexRecord> records)
        {
            object? itemsObject = null;
            object? foldersObject = null;

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

                    if (itemsObject != null)
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
                                    out IndexRecord? existing))
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

                                int processed =
                                    _incrementalNewCount +
                                    _incrementalChangedCount;

                                if (processed > 0 &&
                                    processed % 20 == 0)
                                {
                                    ProgressCountText.Text =
                                        $"{_incrementalNewCount:N0} neu | " +
                                        $"{_incrementalChangedCount:N0} geändert";

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
                }
                catch
                {
                }

                foldersObject =
                    folder.Folders;

                if (foldersObject != null)
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
            int legacyCounter = 0;

            while ((line =
                reader.ReadLine()) != null)
            {
                string[] columns =
                    line.Split(
                        new[] { '\t' },
                        14,
                        StringSplitOptions.None);

                if (columns.Length != 14)
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
                        Date = columns[1],
                        Sender = columns[2],
                        Recipient = columns[3],
                        Cc = columns[4],
                        Mailbox = columns[5],
                        Subject = columns[6],
                        Folder = columns[7],
                        Flag = columns[8],
                        Attachment = columns[9],
                        Categories = columns[10],
                        ConversationId = columns[11],
                        EntryId = columns[12],
                        Body = columns[13]
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

                records[key] =
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
                _indexPath + ".tmp";

            List<IndexRecord> orderedRecords =
                records
                    .OrderByDescending(
                        x => ParseIndexDate(
                            x.Date) ??
                            DateTime.MinValue)
                    .ToList();

            using (
                StreamWriter writer =
                    new StreamWriter(
                        temporaryPath,
                        false,
                        new UTF8Encoding(true)))
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
                    "E-Mail-Text");

                int number = 0;

                foreach (IndexRecord record
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

            File.Copy(
                temporaryPath,
                _indexPath,
                true);

            File.Delete(
                temporaryPath);
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
                $"{CleanIndexText(record.Body)}";
        }

        private static IndexRecord BuildIndexRecord(
            dynamic item,
            string storeName,
            string currentPath)
        {
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
                    HasAttachments(
                        item)
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
                        "Body")
            };
        }

        private static DateTime?
            SafeLastModificationTime(
                dynamic item)
        {
            try
            {
                DateTime value =
                    item.LastModificationTime;

                return value;
            }
            catch
            {
                try
                {
                    DateTime value =
                        item.ReceivedTime;

                    return value;
                }
                catch
                {
                    try
                    {
                        DateTime value =
                            item.SentOn;

                        return value;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }

        // =========================================================
        // ERSTER VOLLSTÄNDIGER INDEX
        // =========================================================

        private async Task CreateFullIndexAsync()
        {
            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            _updateInProgress = true;

            try
            {
                CreateIndexButton.IsEnabled = false;
                ConnectOutlookButton.IsEnabled = false;
                SearchButton.IsEnabled = false;

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                _totalFolderCount = 0;
                _totalMailCount = 0;
                _processedMailCount = 0;

                ProgressPanel.Visibility =
                    Visibility.Visible;

                IndexProgressBar.IsIndeterminate =
                    true;

                ProgressPercentText.Text = "";

                ProgressPhaseText.Text =
                    "Phase 1 von 2 – Outlook-Bestand wird gezählt";

                ProgressCountText.Text =
                    "Outlook-Bestand wird vorbereitet ...";

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
                    object? storeObject = null;
                    object? rootFolderObject = null;

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

                        await RefreshUi();

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject != null)
                        {
                            await CountFolderAsync(
                                rootFolderObject,
                                storeName);
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

                if (_totalMailCount <= 0)
                {
                    throw new InvalidOperationException(
                        "Es wurden keine Outlook-E-Mails gefunden.");
                }

                Dictionary<string, IndexRecord> records =
                    new Dictionary<string, IndexRecord>(
                        StringComparer.OrdinalIgnoreCase);

                ProgressPhaseText.Text =
                    "Phase 2 von 2 – E-Mail-Inhalte werden indiziert";

                ProgressPercentText.Text =
                    "0 %";

                ProgressCountText.Text =
                    $"0 von {_totalMailCount:N0} E-Mails verarbeitet";

                IndexProgressBar.IsIndeterminate =
                    false;

                IndexProgressBar.Minimum = 0;
                IndexProgressBar.Maximum =
                    _totalMailCount;

                IndexProgressBar.Value = 0;

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject = null;
                    object? rootFolderObject = null;

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

                        if (rootFolderObject != null)
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

                IndexProgressBar.Maximum = 100;
                IndexProgressBar.Value = 100;

                ProgressPercentText.Text =
                    "100 %";

                ProgressPhaseText.Text =
                    "Indizierung abgeschlossen";

                ProgressCountText.Text =
                    $"{records.Count:N0} E-Mails erfolgreich indiziert";

                IndexStatusText.Text =
                    "Outlook-Inhaltsindex erfolgreich erstellt";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                IndexDetailsText.Text =
                    $"{storeCount} Postfächer | " +
                    $"{_totalFolderCount:N0} Ordner | " +
                    $"{records.Count:N0} E-Mails";

                SearchButton.IsEnabled = true;

                CreateIndexButton.Content =
                    "Jetzt aktualisieren";

                LoadMailboxesFromIndex();
                UpdateLastSyncDisplay();

                SearchStatusText.Text =
                    "Index bereit. Suche und Filter können verwendet werden.";

                await Task.Delay(
                    700);

                ProgressPanel.Visibility =
                    Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                IndexProgressBar.IsIndeterminate =
                    false;

                IndexStatusText.Text =
                    "Outlook-Indizierung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 40, 40));

                IndexDetailsText.Text =
                    ex.Message;

                ProgressPhaseText.Text =
                    "Indizierung abgebrochen";
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);

                CreateIndexButton.IsEnabled = true;
                ConnectOutlookButton.IsEnabled = true;

                _updateInProgress = false;
            }
        }

        private async Task CountFolderAsync(
            object folderObject,
            string storeName)
        {
            object? itemsObject = null;
            object? foldersObject = null;

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

                    if (itemsObject != null)
                    {
                        dynamic items =
                            itemsObject;

                        int itemCount =
                            items.Count;

                        for (int i = 1;
                             i <= itemCount;
                             i++)
                        {
                            object? itemObject = null;

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
                                        250 == 0)
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

                if (foldersObject != null)
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
            object? itemsObject = null;
            object? foldersObject = null;

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
                                25 == 0)
                            {
                                double percent =
                                    _totalMailCount > 0
                                        ? (double)_processedMailCount /
                                          _totalMailCount *
                                          100
                                        : 0;

                                IndexProgressBar.Value =
                                    Math.Min(
                                        _processedMailCount,
                                        IndexProgressBar.Maximum);

                                ProgressPercentText.Text =
                                    $"{Math.Min(percent, 100):0.0} %";

                                ProgressCountText.Text =
                                    $"{_processedMailCount:N0} von " +
                                    $"{_totalMailCount:N0} E-Mails verarbeitet";

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

                if (foldersObject != null)
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
            if (e.Key == Key.Enter &&
                SearchButton.IsEnabled)
            {
                await ExecuteSearchAsync();
            }
        }

        private void ResetFiltersButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            FromDatePicker.SelectedDate = null;
            ToDatePicker.SelectedDate = null;

            MailboxComboBox.SelectedIndex = 0;
            AttachmentComboBox.SelectedIndex = 0;
            FlagComboBox.SelectedIndex = 0;
            SortComboBox.SelectedIndex = 0;

            SubjectOnlyCheckBox.IsChecked = false;

            SenderFilterTextBox.Text = "";
            RecipientFilterTextBox.Text = "";

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

            string mailbox =
                MailboxComboBox.SelectedItem?.ToString()
                ?? "Alle Postfächer";

            string attachment =
                GetComboBoxText(
                    AttachmentComboBox);

            string flag =
                GetComboBoxText(
                    FlagComboBox);

            string sort =
                GetComboBoxText(
                    SortComboBox);

            bool subjectOnly =
                SubjectOnlyCheckBox.IsChecked ==
                true;

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

            SearchButton.IsEnabled = false;
            SearchTextBox.IsEnabled = false;

            SearchStatusText.Text =
                "Index wird durchsucht ...";

            try
            {
                SearchOptions options =
                    new SearchOptions
                    {
                        Query = query,
                        FromDate = fromDate,
                        ToDate = toDate,
                        Mailbox = mailbox,
                        Attachment = attachment,
                        Flag = flag,
                        SenderFilter = senderFilter,
                        RecipientFilter = recipientFilter,
                        SubjectOnly = subjectOnly,
                        Sort = sort
                    };

                ActiveFiltersText.Text =
                    BuildActiveFiltersText(
                        options);

                SearchResponse response =
                    await Task.Run(
                        () => SearchIndex(
                            options));

                SearchResultsGrid.ItemsSource =
                    response.Results;

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
                SearchButton.IsEnabled = true;
                SearchTextBox.IsEnabled = true;
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
                        ? options.FromDate.Value.ToString(
                            "dd.MM.yyyy")
                        : "offen";

                string to =
                    options.ToDate.HasValue
                        ? options.ToDate.Value.ToString(
                            "dd.MM.yyyy")
                        : "offen";

                filters.Add(
                    $"Zeitraum: {from} bis {to}");
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

            if (options.SubjectOnly)
            {
                filters.Add(
                    "Nur Betreff");
            }

            if (filters.Count == 0)
            {
                return "Aktive Filter: keine";
            }

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
                    StringSplitOptions.RemoveEmptyEntries |
                    StringSplitOptions.TrimEntries);

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
                        new[] { '\t' },
                        14,
                        StringSplitOptions.None);

                if (columns.Length != 14 ||
                    !int.TryParse(
                        columns[0],
                        out _))
                {
                    continue;
                }

                DateTime? mailDate =
                    ParseIndexDate(
                        columns[1]);

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

                string sender = columns[2];
                string recipient = columns[3];
                string cc = columns[4];
                string mailbox = columns[5];
                string subject = columns[6];
                string folder = columns[7];
                string flag = columns[8];
                string attachment = columns[9];
                string categories = columns[10];
                string conversationId = columns[11];
                string entryId = columns[12];
                string body = columns[13];

                // NEU BUILD 1080:
                // Freie Texteingabe für Absender.
                if (!string.IsNullOrWhiteSpace(
                        options.SenderFilter) &&
                    !sender.Contains(
                        options.SenderFilter,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // NEU BUILD 1080:
                // Freie Texteingabe für Empfänger.
                //
                // Wir prüfen Empfänger UND CC.
                // Damit wird eine Person auch gefunden,
                // wenn sie nur in Kopie angeschrieben wurde.
                if (!string.IsNullOrWhiteSpace(
                        options.RecipientFilter) &&
                    !recipient.Contains(
                        options.RecipientFilter,
                        StringComparison.OrdinalIgnoreCase) &&
                    !cc.Contains(
                        options.RecipientFilter,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (options.Mailbox !=
                    "Alle Postfächer" &&
                    !mailbox.Equals(
                        options.Mailbox,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                bool hasAttachment =
                    attachment == "1";

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
                    !flag.Equals(
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
                    !flag.Contains(
                        "Erledigt",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (options.Flag ==
                        "Fähnchen" &&
                    !flag.Contains(
                        "Fähnchen",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (searchWords.Length > 0)
                {
                    string searchableText =
                        options.SubjectOnly
                            ? subject
                            : sender + " " +
                              recipient + " " +
                              cc + " " +
                              mailbox + " " +
                              subject + " " +
                              folder + " " +
                              flag + " " +
                              categories + " " +
                              body;

                    bool allWordsFound =
                        searchWords.All(
                            word =>
                                searchableText.Contains(
                                    word,
                                    StringComparison.OrdinalIgnoreCase));

                    if (!allWordsFound)
                    {
                        continue;
                    }
                }

                matches.Add(
                    new SearchResult
                    {
                        SortDate =
                            mailDate ??
                            DateTime.MinValue,

                        Date =
                            mailDate.HasValue
                                ? mailDate.Value.ToString(
                                    "dd.MM.yyyy HH:mm")
                                : "",

                        Mailbox = mailbox,
                        Recipient = recipient,
                        Cc = cc,
                        Sender = sender,
                        Flag = flag,

                        Attachment =
                            hasAttachment
                                ? "Ja"
                                : "Nein",

                        Subject = subject,
                        Folder = folder,
                        Categories = categories,
                        ConversationId = conversationId,
                        EntryId = entryId,
                        Body = body
                    });
            }

            IEnumerable<SearchResult> sorted =
                options.Sort ==
                    "Älteste zuerst"
                    ? matches.OrderBy(
                        x => x.SortDate)
                    : matches.OrderByDescending(
                        x => x.SortDate);

            int totalMatches =
                matches.Count;

            List<SearchResult> displayed =
                sorted
                    .Take(
                        MaximumSearchResults)
                    .ToList();

            return new SearchResponse
            {
                Results = displayed,
                TotalMatches = totalMatches,
                WasLimited =
                    totalMatches >
                    MaximumSearchResults
            };
        }

        // =========================================================
        // ORIGINAL-MAIL AUS OUTLOOK ÖFFNEN
        // =========================================================

        private void SearchResultsGrid_MouseDoubleClick(
            object sender,
            MouseButtonEventArgs e)
        {
            if (SearchResultsGrid.SelectedItem
                is not SearchResult result)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(
                result.EntryId))
            {
                MessageBox.Show(
                    "Für diesen Treffer ist keine Outlook-EntryID gespeichert.",
                    "Original-Mail öffnen",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

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

                if (outlookNamespace == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-MAPI-Schnittstelle konnte nicht geöffnet werden.");
                }

                dynamic outlookNs =
                    outlookNamespace;

                stores =
                    outlookNs.Stores;

                if (stores != null &&
                    !string.IsNullOrWhiteSpace(
                        result.Mailbox))
                {
                    dynamic outlookStores =
                        stores;

                    int storeCount =
                        outlookStores.Count;

                    for (int i = 1;
                         i <= storeCount;
                         i++)
                    {
                        object? storeObject = null;

                        try
                        {
                            storeObject =
                                outlookStores.Item(i);

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
                                        outlookNs.GetItemFromID(
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
                            outlookNs.GetItemFromID(
                                result.EntryId);
                    }
                    catch
                    {
                        mailItem = null;
                    }
                }

                if (mailItem == null)
                {
                    MessageBox.Show(
                        "Die Original-Mail konnte in Outlook nicht mehr gefunden werden.\n\n" +
                        "Möglicherweise wurde sie seit der letzten Index-Aktualisierung verschoben oder gelöscht.\n" +
                        "Bitte den Suchindex aktualisieren und erneut versuchen.",
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

        // =========================================================
        // OUTLOOK-HILFSMETHODEN
        // =========================================================

        private static bool HasAttachments(
            dynamic item)
        {
            object? attachmentsObject =
                null;

            try
            {
                attachmentsObject =
                    item.Attachments;

                if (attachmentsObject ==
                    null)
                {
                    return false;
                }

                dynamic attachments =
                    attachmentsObject;

                return attachments.Count >
                       0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(
                    attachmentsObject);
            }
        }

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
                    return string.IsNullOrWhiteSpace(
                        flagRequest)
                        ? "Erledigt / Häkchen"
                        : "Erledigt / Häkchen – " +
                          flagRequest;
                }

                if (flagStatus == 2)
                {
                    return string.IsNullOrWhiteSpace(
                        flagRequest)
                        ? "Fähnchen"
                        : "Fähnchen – " +
                          flagRequest;
                }

                return "Keine Kennzeichnung";
            }
            catch
            {
                return "Keine Kennzeichnung";
            }
        }

        private static string SafeDynamicString(
            dynamic item,
            string propertyName)
        {
            try
            {
                switch (propertyName)
                {
                    case "Subject":
                        return item.Subject ??
                               "";

                    case "SenderName":
                        return item.SenderName ??
                               "";

                    case "Body":
                        return item.Body ??
                               "";

                    case "DisplayName":
                        return item.DisplayName ??
                               "";

                    case "Name":
                        return item.Name ??
                               "";

                    case "To":
                        return item.To ??
                               "";

                    case "CC":
                        return item.CC ??
                               "";

                    case "Categories":
                        return item.Categories ??
                               "";

                    case "ConversationID":
                        return item.ConversationID ??
                               "";

                    case "EntryID":
                        return item.EntryID ??
                               "";
                }
            }
            catch
            {
            }

            return "";
        }

        private static string SafeDynamicDateTime(
            dynamic item)
        {
            try
            {
                DateTime value =
                    item.ReceivedTime;

                if (value.Year >
                    1900)
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
                return item.Content
                    ?.ToString() ??
                    "";
            }

            return comboBox.SelectedItem
                ?.ToString() ??
                "";
        }

        private static string CleanIndexText(
            string? text)
        {
            if (string.IsNullOrWhiteSpace(
                text))
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
                1024L * 1024L * 1024L)
            {
                return
                    $"{bytes / (1024.0 * 1024.0 * 1024.0):0.0} GB";
            }

            if (bytes >=
                1024L * 1024L)
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
                    () => { },
                    DispatcherPriority.Background);
        }

        private static void ReleaseComObject(
            object? comObject)
        {
            if (comObject != null &&
                Marshal.IsComObject(
                    comObject))
            {
                Marshal.FinalReleaseComObject(
                    comObject);
            }
        }

        // =========================================================
        // DATENKLASSEN
        // =========================================================

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
                    Body == other.Body;
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

            // Neu in Build 1080
            public string SenderFilter { get; set; } = "";
            public string RecipientFilter { get; set; } = "";

            public string Sort { get; set; } =
                "Neueste zuerst";

            public bool SubjectOnly { get; set; }
        }

        private class SearchResponse
        {
            public List<SearchResult> Results
            {
                get;
                set;
            } =
                new List<SearchResult>();

            public int TotalMatches
            {
                get;
                set;
            }

            public bool WasLimited
            {
                get;
                set;
            }
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
            public string Subject { get; set; } = "";
            public string Folder { get; set; } = "";
            public string Categories { get; set; } = "";
            public string ConversationId { get; set; } = "";
            public string EntryId { get; set; } = "";
            public string Body { get; set; } = "";
        }
    }
}