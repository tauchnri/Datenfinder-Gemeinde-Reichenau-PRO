using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace Datenfinder.UI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ConnectOutlookButton_Click(object sender, RoutedEventArgs e)
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
                    Type.GetTypeFromProgID("Outlook.Application");

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

                dynamic outlook = outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace("MAPI");

                if (outlookNamespace == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-MAPI-Schnittstelle konnte nicht geöffnet werden.");
                }

                dynamic outlookNs = outlookNamespace;

                stores = outlookNs.Stores;

                if (stores == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-Datenspeicher konnten nicht gelesen werden.");
                }

                dynamic outlookStores = stores;

                int storeCount = outlookStores.Count;

                OutlookStatusText.Text =
                    "Status: Outlook erfolgreich verbunden";

                OutlookStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                OutlookDetailsText.Text =
                    $"Gefundene Outlook-Datenspeicher: {storeCount}";

                CreateIndexButton.IsEnabled = true;

                IndexStatusText.Text =
                    "Outlook erkannt – Index noch nicht erstellt";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                IndexDetailsText.Text =
                    "Klicken Sie jetzt auf „Index erstellen“, um die Outlook-Ordner auszuwerten.";
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

                CreateIndexButton.IsEnabled = false;

                IndexStatusText.Text =
                    "Suchindex noch nicht erstellt";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(119, 119, 119));

                IndexDetailsText.Text =
                    "Klicken Sie zuerst auf „Outlook prüfen“.";
            }
            finally
            {
                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private void CreateIndexButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            object? outlookApplication = null;
            object? outlookNamespace = null;
            object? stores = null;

            try
            {
                CreateIndexButton.IsEnabled = false;

                IndexStatusText.Text =
                    "Outlook-Ordner werden ausgewertet – bitte warten ...";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(85, 85, 85));

                IndexDetailsText.Text =
                    "Ordner und E-Mails werden gelesen und protokolliert.";

                Type? outlookType =
                    Type.GetTypeFromProgID("Outlook.Application");

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

                dynamic outlook = outlookApplication;

                outlookNamespace =
                    outlook.GetNamespace("MAPI");

                if (outlookNamespace == null)
                {
                    throw new InvalidOperationException(
                        "Die Outlook-MAPI-Schnittstelle konnte nicht geöffnet werden.");
                }

                dynamic outlookNs = outlookNamespace;

                stores = outlookNs.Stores;

                dynamic outlookStores = stores;

                int storeCount = outlookStores.Count;
                int folderCount = 0;
                int mailCount = 0;

                StringBuilder report = new StringBuilder();

                report.AppendLine("Datenfinder Gemeinde Reichenau PRO");
                report.AppendLine("Outlook-Ordnerbericht");
                report.AppendLine(
                    $"Erstellt am: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                report.AppendLine();
                report.AppendLine(
                    $"Gefundene Datenspeicher: {storeCount}");
                report.AppendLine();

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

                        dynamic store = storeObject;

                        string storeName = "Unbekannter Datenspeicher";

                        try
                        {
                            storeName = store.DisplayName;
                        }
                        catch
                        {
                        }

                        report.AppendLine(
                            $"DATENSPEICHER: {storeName}");
                        report.AppendLine(
                            new string('=', 70));

                        rootFolderObject =
                            store.GetRootFolder();

                        if (rootFolderObject != null)
                        {
                            CountFolderContents(
                                rootFolderObject,
                                ref folderCount,
                                ref mailCount,
                                report,
                                0);
                        }

                        report.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        report.AppendLine(
                            $"Fehler beim Datenspeicher: {ex.Message}");
                        report.AppendLine();
                    }
                    finally
                    {
                        ReleaseComObject(rootFolderObject);
                        ReleaseComObject(storeObject);
                    }
                }

                string reportFolder =
                    Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments),
                        "Datenfinder Gemeinde Reichenau PRO");

                Directory.CreateDirectory(reportFolder);

                string reportPath =
                    Path.Combine(
                        reportFolder,
                        "Outlook-Ordnerbericht.txt");

                report.AppendLine();
                report.AppendLine(
                    new string('=', 70));
                report.AppendLine(
                    $"Datenspeicher gesamt: {storeCount}");
                report.AppendLine(
                    $"Ordner gesamt: {folderCount}");
                report.AppendLine(
                    $"E-Mails gesamt: {mailCount}");

                File.WriteAllText(
                    reportPath,
                    report.ToString(),
                    Encoding.UTF8);

                IndexStatusText.Text =
                    "Outlook-Auswertung abgeschlossen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(0, 120, 70));

                IndexDetailsText.Text =
                    $"Datenspeicher: {storeCount} | Ordner: {folderCount} | E-Mails: {mailCount}\n" +
                    $"Ordnerbericht gespeichert unter: {reportPath}";

                MessageBox.Show(
                    "Die Outlook-Auswertung ist abgeschlossen.\n\n" +
                    $"Ordner: {folderCount}\n" +
                    $"E-Mails: {mailCount}\n\n" +
                    "Der vollständige Ordnerbericht wurde gespeichert unter:\n" +
                    reportPath,
                    "Datenfinder Gemeinde Reichenau PRO",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                IndexStatusText.Text =
                    "Outlook-Auswertung fehlgeschlagen";

                IndexStatusText.Foreground =
                    new SolidColorBrush(
                        Color.FromRgb(180, 40, 40));

                IndexDetailsText.Text =
                    ex.Message;
            }
            finally
            {
                CreateIndexButton.IsEnabled = true;

                ReleaseComObject(stores);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private static void CountFolderContents(
            object folderObject,
            ref int folderCount,
            ref int mailCount,
            StringBuilder report,
            int level)
        {
            object? itemsObject = null;
            object? foldersObject = null;

            try
            {
                dynamic folder = folderObject;

                folderCount++;

                string folderName = "Unbekannter Ordner";

                try
                {
                    folderName = folder.Name;
                }
                catch
                {
                }

                int mailsInThisFolder = 0;

                try
                {
                    itemsObject = folder.Items;

                    if (itemsObject != null)
                    {
                        dynamic items = itemsObject;

                        int itemCount = items.Count;

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

                                dynamic item = itemObject;

                                try
                                {
                                    int itemClass = item.Class;

                                    // Outlook MailItem = 43
                                    if (itemClass == 43)
                                    {
                                        mailCount++;
                                        mailsInThisFolder++;
                                    }
                                }
                                catch
                                {
                                }
                            }
                            catch
                            {
                            }
                            finally
                            {
                                ReleaseComObject(itemObject);
                            }
                        }
                    }
                }
                catch
                {
                }

                string indent =
                    new string(' ', level * 2);

                report.AppendLine(
                    $"{indent}- {folderName} | E-Mails: {mailsInThisFolder}");

                foldersObject = folder.Folders;

                if (foldersObject != null)
                {
                    dynamic folders = foldersObject;

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
                                CountFolderContents(
                                    subFolderObject,
                                    ref folderCount,
                                    ref mailCount,
                                    report,
                                    level + 1);
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
    }
}