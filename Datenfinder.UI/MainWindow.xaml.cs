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
        private const string IndexSchema = "1050";

        private int _totalFolderCount;
        private int _totalMailCount;
        private int _processedMailCount;

        private readonly string _indexFolder;
        private readonly string _indexPath;

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

            InitializeFilters();
            CheckExistingIndex();
        }

        private void InitializeFilters()
        {
            AttachmentComboBox.SelectedIndex = 0;
            FlagComboBox.SelectedIndex = 0;
            SortComboBox.SelectedIndex = 0;

            MailboxComboBox.Items.Clear();
            MailboxComboBox.Items.Add("Alle Postfächer");
            MailboxComboBox.SelectedIndex = 0;
        }

        private void CheckExistingIndex()
        {
            if (!File.Exists(_indexPath))
            {
                SearchButton.IsEnabled = false;
                return;
            }

            if (!IsCurrentIndex())
            {
                SearchButton.IsEnabled = false;

                IndexStatusText.Text =
                    "Vorhandener Suchindex ist veraltet";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 100, 0));

                IndexDetailsText.Text =
                    "Für Build 1050 muss der Outlook-Index einmal neu erstellt werden.";

                SearchStatusText.Text =
                    "Bitte Outlook prüfen und anschließend den Index neu erstellen.";

                return;
            }

            SearchButton.IsEnabled = true;

            IndexStatusText.Text =
                "Outlook-Inhaltsindex ist bereit";

            IndexStatusText.Foreground =
                new SolidColorBrush(
                    Color.FromRgb(0, 120, 70));

            FileInfo fileInfo =
                new FileInfo(_indexPath);

            IndexDetailsText.Text =
                $"Indexgröße: {FormatFileSize(fileInfo.Length)}";

            SearchStatusText.Text =
                "Suchbegriff eingeben oder die Filter verwenden.";

            LoadMailboxesFromIndex();
        }

        private bool IsCurrentIndex()
        {
            try
            {
                using StreamReader reader =
                    new StreamReader(
                        _indexPath,
                        Encoding.UTF8,
                        true);

                for (int i = 0; i < 12; i++)
                {
                    string? line = reader.ReadLine();

                    if (line == null)
                    {
                        break;
                    }

                    if (line.Trim() ==
                        $"Schema: {IndexSchema}")
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
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

                    if (!int.TryParse(columns[0], out _))
                    {
                        continue;
                    }

                    string mailbox =
                        columns[5].Trim();

                    if (!string.IsNullOrWhiteSpace(mailbox))
                    {
                        mailboxes.Add(mailbox);
                    }
                }

                string? selected =
                    MailboxComboBox.SelectedItem?.ToString();

                MailboxComboBox.Items.Clear();
                MailboxComboBox.Items.Add("Alle Postfächer");

                foreach (string mailbox
                    in mailboxes.OrderBy(x => x))
                {
                    MailboxComboBox.Items.Add(mailbox);
                }

                if (!string.IsNullOrWhiteSpace(selected) &&
                    MailboxComboBox.Items.Contains(selected))
                {
                    MailboxComboBox.SelectedItem = selected;
                }
                else
                {
                    MailboxComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                MailboxComboBox.Items.Clear();
                MailboxComboBox.Items.Add("Alle Postfächer");
                MailboxComboBox.SelectedIndex = 0;
            }
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

                OutlookStatusText.Text =
                    "Status: Outlook erfolgreich verbunden";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                OutlookDetailsText.Text =
                    $"Gefundene Datenspeicher/Postfächer: {storeCount}";

                CreateIndexButton.IsEnabled = true;

                if (IsCurrentIndex())
                {
                    IndexStatusText.Text =
                        "Index vorhanden – kann bei Bedarf aktualisiert werden";

                    IndexDetailsText.Text =
                        "Der aktuelle Build-1050-Index kann bereits durchsucht werden.";

                    SearchButton.IsEnabled = true;
                }
                else
                {
                    IndexStatusText.Text =
                        "Build-1050-Index muss erstellt werden";

                    IndexDetailsText.Text =
                        "Der Index enthält Postfach, Empfänger, Kennzeichnung und Anhang.";

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

                OutlookDetailsText.Text = ex.Message;

                CreateIndexButton.IsEnabled = false;
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
            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            try
            {
                CreateIndexButton.IsEnabled = false;
                ConnectOutlookButton.IsEnabled = false;
                SearchButton.IsEnabled = false;

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                SearchStatusText.Text = "";

                _totalFolderCount = 0;
                _totalMailCount = 0;
                _processedMailCount = 0;

                ProgressPanel.Visibility =
                    Visibility.Visible;

                IndexProgressBar.IsIndeterminate = true;
                IndexProgressBar.Minimum = 0;
                IndexProgressBar.Maximum = 100;
                IndexProgressBar.Value = 0;

                ProgressPercentText.Text = "";
                ProgressCountText.Text =
                    "Outlook-Bestand wird vorbereitet ...";

                ProgressFolderText.Text = "";

                IndexStatusText.Text =
                    "E-Mails werden indiziert – bitte warten ...";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(85, 85, 85));

                Type? outlookType =
                    Type.GetTypeFromProgID(
                        "Outlook.Application");

                if (outlookType == null)
                {
                    throw new InvalidOperationException(
                        "Das klassische Microsoft Outlook wurde nicht gefunden.");
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

                if (outlookNamespace == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-MAPI-Schnittstelle konnte nicht geöffnet werden.");
                }

                dynamic outlookNs =
                    outlookNamespace;

                stores =
                    outlookNs.Stores;

                dynamic outlookStores =
                    stores;

                int storeCount =
                    outlookStores.Count;

                // PHASE 1
                ProgressPhaseText.Text =
                    "Phase 1 von 2 – Outlook-Bestand wird gezählt";

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject = null;
                    object? rootFolderObject = null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

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
                        ReleaseComObject(rootFolderObject);
                        ReleaseComObject(storeObject);
                    }
                }

                if (_totalMailCount <= 0)
                {
                    throw new InvalidOperationException(
                        "Es wurden keine Outlook-E-Mails gefunden.");
                }

                Directory.CreateDirectory(
                    _indexFolder);

                // PHASE 2
                ProgressPhaseText.Text =
                    "Phase 2 von 2 – E-Mail-Inhalte und Zusatzdaten werden indiziert";

                IndexProgressBar.IsIndeterminate = false;
                IndexProgressBar.Minimum = 0;
                IndexProgressBar.Maximum =
                    _totalMailCount;

                IndexProgressBar.Value = 0;

                ProgressPercentText.Text = "0 %";

                ProgressCountText.Text =
                    $"0 von {_totalMailCount:N0} E-Mails verarbeitet";

                using StreamWriter writer =
                    new StreamWriter(
                        _indexPath,
                        false,
                        new UTF8Encoding(true));

                await writer.WriteLineAsync(
                    "Datenfinder Gemeinde Reichenau PRO");

                await writer.WriteLineAsync(
                    "Outlook-Inhaltsindex");

                await writer.WriteLineAsync(
                    $"Schema: {IndexSchema}");

                await writer.WriteLineAsync(
                    $"Erstellt am: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");

                await writer.WriteLineAsync(
                    $"Postfächer: {storeCount}");

                await writer.WriteLineAsync(
                    $"Ordner: {_totalFolderCount}");

                await writer.WriteLineAsync(
                    $"E-Mails bei Start: {_totalMailCount}");

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

                for (int storeIndex = 1;
                     storeIndex <= storeCount;
                     storeIndex++)
                {
                    object? storeObject = null;
                    object? rootFolderObject = null;

                    try
                    {
                        storeObject =
                            outlookStores.Item(storeIndex);

                        dynamic store =
                            storeObject;

                        string storeName =
                            SafeDynamicString(
                                store,
                                "DisplayName");

                        if (string.IsNullOrWhiteSpace(storeName))
                        {
                            storeName =
                                "Unbekanntes Postfach";
                        }

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject != null)
                        {
                            await IndexFolderAsync(
                                rootFolderObject,
                                storeName,
                                "",
                                writer);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(rootFolderObject);
                        ReleaseComObject(storeObject);
                    }
                }

                await writer.FlushAsync();

                int finalTotal =
                    Math.Max(
                        _totalMailCount,
                        _processedMailCount);

                IndexProgressBar.Maximum =
                    finalTotal;

                IndexProgressBar.Value =
                    Math.Min(
                        _processedMailCount,
                        finalTotal);

                ProgressPercentText.Text = "100 %";

                ProgressCountText.Text =
                    $"{_processedMailCount:N0} E-Mails erfolgreich indiziert";

                ProgressPhaseText.Text =
                    "Indizierung abgeschlossen";

                ProgressFolderText.Text =
                    $"Index gespeichert unter: {_indexPath}";

                IndexStatusText.Text =
                    "Outlook-Inhaltsindex Build 1050 erfolgreich erstellt";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                IndexDetailsText.Text =
                    $"{storeCount} Postfächer | " +
                    $"{_totalFolderCount:N0} Ordner | " +
                    $"{_processedMailCount:N0} E-Mails";

                SearchButton.IsEnabled = true;

                SearchStatusText.Text =
                    "Index bereit. Suche und Filter können verwendet werden.";

                LoadMailboxesFromIndex();

                await Task.Delay(600);

                ProgressPanel.Visibility =
                    Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                IndexProgressBar.IsIndeterminate = false;

                IndexStatusText.Text =
                    "Outlook-Indizierung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 40, 40));

                IndexDetailsText.Text =
                    ex.Message;

                ProgressPhaseText.Text =
                    "Indizierung abgebrochen";

                SearchButton.IsEnabled =
                    IsCurrentIndex();
            }
            finally
            {
                CreateIndexButton.IsEnabled = true;
                ConnectOutlookButton.IsEnabled = true;

                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

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

            SearchTextBox.Focus();
        }

        private async Task ExecuteSearchAsync()
        {
            if (!File.Exists(_indexPath) ||
                !IsCurrentIndex())
            {
                SearchStatusText.Text =
                    "Der Build-1050-Index wurde noch nicht erstellt.";

                SearchResultsGrid.Visibility =
                    Visibility.Collapsed;

                return;
            }

            string query =
                SearchTextBox.Text.Trim();

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
                SubjectOnlyCheckBox.IsChecked == true;

            bool noCriteria =
                string.IsNullOrWhiteSpace(query) &&
                !fromDate.HasValue &&
                !toDate.HasValue &&
                mailbox == "Alle Postfächer" &&
                attachment == "Alle" &&
                flag == "Alle";

            if (noCriteria)
            {
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
                        SubjectOnly = subjectOnly,
                        Sort = sort
                    };

                SearchResponse response =
                    await Task.Run(
                        () => SearchIndex(options));

                SearchResultsGrid.ItemsSource =
                    response.Results;

                if (response.Results.Count == 0)
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

        private SearchResponse SearchIndex(
            SearchOptions options)
        {
            List<SearchResult> allMatches =
                new List<SearchResult>();

            string[] searchWords =
                options.Query
                    .Split(
                        new[] { ' ' },
                        StringSplitOptions.RemoveEmptyEntries |
                        StringSplitOptions.TrimEntries);

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

                DateTime? mailDate =
                    ParseIndexDate(
                        columns[1]);

                if (options.FromDate.HasValue)
                {
                    if (!mailDate.HasValue ||
                        mailDate.Value.Date <
                        options.FromDate.Value.Date)
                    {
                        continue;
                    }
                }

                if (options.ToDate.HasValue)
                {
                    if (!mailDate.HasValue ||
                        mailDate.Value.Date >
                        options.ToDate.Value.Date)
                    {
                        continue;
                    }
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
                    string searchableText;

                    if (options.SubjectOnly)
                    {
                        searchableText = subject;
                    }
                    else
                    {
                        searchableText =
                            sender + " " +
                            recipient + " " +
                            cc + " " +
                            mailbox + " " +
                            subject + " " +
                            folder + " " +
                            flag + " " +
                            categories + " " +
                            body;
                    }

                    bool allWordsFound = true;

                    foreach (string word
                        in searchWords)
                    {
                        if (!searchableText.Contains(
                            word,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            allWordsFound = false;
                            break;
                        }
                    }

                    if (!allWordsFound)
                    {
                        continue;
                    }
                }

                allMatches.Add(
                    new SearchResult
                    {
                        SortDate =
                            mailDate ?? DateTime.MinValue,

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
                        ConversationId =
                            conversationId,
                        EntryId = entryId,
                        Body = body
                    });
            }

            IEnumerable<SearchResult> sorted;

            if (options.Sort ==
                "Älteste zuerst")
            {
                sorted =
                    allMatches.OrderBy(
                        x => x.SortDate);
            }
            else
            {
                sorted =
                    allMatches.OrderByDescending(
                        x => x.SortDate);
            }

            int totalMatches =
                allMatches.Count;

            List<SearchResult> displayed =
                sorted
                    .Take(MaximumSearchResults)
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

                if (string.IsNullOrWhiteSpace(folderName))
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

                                if (itemObject == null)
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
                        object? subFolderObject = null;

                        try
                        {
                            subFolderObject =
                                folders.Item(i);

                            if (subFolderObject != null)
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
                ReleaseComObject(itemsObject);
                ReleaseComObject(foldersObject);
            }
        }

        private async Task IndexFolderAsync(
            object folderObject,
            string storeName,
            string parentPath,
            StreamWriter writer)
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

                if (string.IsNullOrWhiteSpace(folderName))
                {
                    folderName =
                        "Unbekannter Ordner";
                }

                string currentPath =
                    string.IsNullOrWhiteSpace(parentPath)
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

                                if (itemObject == null)
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

                                string date =
                                    SafeDynamicDateTime(
                                        item);

                                string sender =
                                    SafeDynamicString(
                                        item,
                                        "SenderName");

                                string recipient =
                                    SafeDynamicString(
                                        item,
                                        "To");

                                string cc =
                                    SafeDynamicString(
                                        item,
                                        "CC");

                                string subject =
                                    SafeDynamicString(
                                        item,
                                        "Subject");

                                string categories =
                                    SafeDynamicString(
                                        item,
                                        "Categories");

                                string conversationId =
                                    SafeDynamicString(
                                        item,
                                        "ConversationID");

                                string entryId =
                                    SafeDynamicString(
                                        item,
                                        "EntryID");

                                string body =
                                    SafeDynamicString(
                                        item,
                                        "Body");

                                string flag =
                                    GetFlagDescription(
                                        item);

                                bool hasAttachment =
                                    HasAttachments(
                                        item);

                                _processedMailCount++;

                                string indexLine =
                                    $"{_processedMailCount}\t" +
                                    $"{CleanIndexText(date)}\t" +
                                    $"{CleanIndexText(sender)}\t" +
                                    $"{CleanIndexText(recipient)}\t" +
                                    $"{CleanIndexText(cc)}\t" +
                                    $"{CleanIndexText(storeName)}\t" +
                                    $"{CleanIndexText(subject)}\t" +
                                    $"{CleanIndexText(currentPath)}\t" +
                                    $"{CleanIndexText(flag)}\t" +
                                    $"{(hasAttachment ? "1" : "0")}\t" +
                                    $"{CleanIndexText(categories)}\t" +
                                    $"{CleanIndexText(conversationId)}\t" +
                                    $"{CleanIndexText(entryId)}\t" +
                                    $"{CleanIndexText(body)}";

                                await writer.WriteLineAsync(
                                    indexLine);

                                int displayTotal =
                                    Math.Max(
                                        _totalMailCount,
                                        _processedMailCount);

                                if (displayTotal >
                                    IndexProgressBar.Maximum)
                                {
                                    IndexProgressBar.Maximum =
                                        displayTotal;
                                }

                                if (_processedMailCount %
                                        25 == 0)
                                {
                                    double percent =
                                        displayTotal > 0
                                            ? (double)_processedMailCount /
                                              displayTotal *
                                              100.0
                                            : 0;

                                    IndexProgressBar.Value =
                                        Math.Min(
                                            _processedMailCount,
                                            IndexProgressBar.Maximum);

                                    ProgressPercentText.Text =
                                        $"{Math.Min(percent, 100):0.0} %";

                                    ProgressCountText.Text =
                                        $"{_processedMailCount:N0} von " +
                                        $"{displayTotal:N0} E-Mails verarbeitet";

                                    await RefreshUi();
                                }
                            }
                            catch
                            {
                                // Einzelne fehlerhafte Mail
                                // soll den gesamten Lauf nicht stoppen.
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
                        object? subFolderObject = null;

                        try
                        {
                            subFolderObject =
                                folders.Item(i);

                            if (subFolderObject != null)
                            {
                                await IndexFolderAsync(
                                    subFolderObject,
                                    storeName,
                                    currentPath,
                                    writer);
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
                ReleaseComObject(itemsObject);
                ReleaseComObject(foldersObject);
            }
        }

        private static bool HasAttachments(
            dynamic item)
        {
            object? attachmentsObject = null;

            try
            {
                attachmentsObject =
                    item.Attachments;

                if (attachmentsObject == null)
                {
                    return false;
                }

                dynamic attachments =
                    attachmentsObject;

                return attachments.Count > 0;
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
                        item.FlagRequest ?? "";
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
                        return item.Subject ?? "";

                    case "SenderName":
                        return item.SenderName ?? "";

                    case "Body":
                        return item.Body ?? "";

                    case "DisplayName":
                        return item.DisplayName ?? "";

                    case "Name":
                        return item.Name ?? "";

                    case "To":
                        return item.To ?? "";

                    case "CC":
                        return item.CC ?? "";

                    case "Categories":
                        return item.Categories ?? "";

                    case "ConversationID":
                        return item.ConversationID ?? "";

                    case "EntryID":
                        return item.EntryID ?? "";
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

                return value.ToString(
                    "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture);
            }
            catch
            {
                return "";
            }
        }

        private static DateTime? ParseIndexDate(
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
                return item.Content?.ToString()
                       ?? "";
            }

            return comboBox.SelectedItem?.ToString()
                   ?? "";
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

            return $"{bytes} Byte";
        }

        private static async Task RefreshUi()
        {
            await Application.Current.Dispatcher.InvokeAsync(
                () => { },
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
            } = new List<SearchResult>();

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
            public DateTime SortDate
            {
                get;
                set;
            }

            public string Date
            {
                get;
                set;
            } = "";

            public string Mailbox
            {
                get;
                set;
            } = "";

            public string Recipient
            {
                get;
                set;
            } = "";

            public string Cc
            {
                get;
                set;
            } = "";

            public string Sender
            {
                get;
                set;
            } = "";

            public string Flag
            {
                get;
                set;
            } = "";

            public string Attachment
            {
                get;
                set;
            } = "";

            public string Subject
            {
                get;
                set;
            } = "";

            public string Folder
            {
                get;
                set;
            } = "";

            public string Categories
            {
                get;
                set;
            } = "";

            public string ConversationId
            {
                get;
                set;
            } = "";

            public string EntryId
            {
                get;
                set;
            } = "";

            public string Body
            {
                get;
                set;
            } = "";
        }
    }
}