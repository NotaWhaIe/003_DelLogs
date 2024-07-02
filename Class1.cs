using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;

namespace DelLogs
{
    public class Class1 : IExternalApplication
    {


        private Timer logTimer;
        private List<Tuple<string, string, int, string>> deletionLogBuffer = new List<Tuple<string, string, int, string>>();
        private readonly string sharedFolderPath = @"\\sb-sharegp\Bim2.0\5. Скрипты\999. BIM-отдел\RevitAutomation\DeleteLog";
        private readonly string usersFolder = @"\\sb-sharegp\Bim2.0\5. Скрипты\999. BIM-отдел\RevitAutomation\DeleteLog";
        private readonly object syncLock = new object();
        private Document currentDocument;
        private bool elementsDeleted = false;
        private string userPrefix = string.Empty;





        public Result OnStartup(UIControlledApplication application)
        {
            if (!IsUserAllowed()) return Result.Cancelled;

            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentChanged += HandleDocumentChanged;
            application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronizedWithCentral;
            application.ControlledApplication.DocumentSaved += OnDocumentSaved;
            SetupTimer();
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            logTimer.Stop();
            logTimer.Dispose();
            if (elementsDeleted)
            {
                SaveAndSync();
                MatchElements();
            }
            return Result.Succeeded;
        }

        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements();
        }

        private void OnDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements();
            if (elementsDeleted)
            {
                SaveAndSync();
                MatchElements();
            }
        }

        private void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            currentDocument = args.Document;
            ExportProjectElements();
            if (elementsDeleted)
            {
                SaveAndSync();
                MatchElements();
            }
        }

        private bool IsUserAllowed()
        {
            var userName = Environment.UserName.ToLowerInvariant();
            var computerName = Environment.MachineName.ToLowerInvariant();
            var userFiles = Directory.GetFiles(usersFolder, "*_users.txt");

            foreach (var userFile in userFiles)
            {
                var users = File.ReadAllLines(userFile, Encoding.UTF8)
                                .Select(u => u.Trim().ToLowerInvariant())
                                .ToList();
                if (users.Contains(userName) || users.Contains(computerName))
                {
                    userPrefix = Path.GetFileNameWithoutExtension(userFile).Split('_')[0];
                    return true;
                }
            }
            return false;
        }

        private void HandleDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            Document doc = args.GetDocument();
            string projectName = Path.GetFileNameWithoutExtension(doc.PathName);
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string userName = doc.Application.Username;

            lock (syncLock)
            {
                foreach (ElementId id in args.GetDeletedElementIds())
                {
                    deletionLogBuffer.Add(new Tuple<string, string, int, string>(projectName, timeStamp, id.IntegerValue, userName));
                    elementsDeleted = true;
                }
            }
        }

        private void SetupTimer()
        {
            logTimer = new Timer(60000) { AutoReset = true, Enabled = true };
            logTimer.Elapsed += (sender, e) =>
            {
                if (elementsDeleted)
                {
                    SaveAndSync();
                    MatchElements();
                }
            };
        }

        private void SaveAndSync()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(sharedFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");
            Directory.CreateDirectory(folderPath);
            Directory.CreateDirectory(sourceTablesFolder);

            string filePath = Path.Combine(sourceTablesFolder, $"{folderName}.csv");

            List<Tuple<string, string, int, string>> bufferCopy;

            lock (syncLock)
            {
                bufferCopy = new List<Tuple<string, string, int, string>>(deletionLogBuffer);
                deletionLogBuffer.Clear();
            }

            TrySaveToFile(() =>
            {
                if (!File.Exists(filePath))
                {
                    File.WriteAllText(filePath, "Project Name,Element ID,Time,User\n", Encoding.UTF8);
                }

                if (bufferCopy.Count > 0)
                {
                    File.AppendAllLines(filePath, bufferCopy.Select(entry =>
                        $"{entry.Item1},{entry.Item3},{entry.Item2},{entry.Item4}"), Encoding.UTF8);
                }
            }, filePath);
        }

        private void MatchElements()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(sharedFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");

            string filePathAR = Path.Combine(sourceTablesFolder, $"{folderName}.csv");
            string filePathDB = Path.Combine(sourceTablesFolder, $"{folderName}_Db.csv");
            string filePathARDeleted = Path.Combine(folderPath, $"{folderName}Deleted.csv");

            TrySaveToFile(() =>
            {
                if (!File.Exists(filePathAR) || !File.Exists(filePathDB)) return;

                var arData = File.ReadAllLines(filePathAR, Encoding.UTF8)
                    .Skip(1)
                    .Select(line => line.Split(','))
                    .GroupBy(parts => int.Parse(parts[1]))
                    .ToDictionary(g => g.Key, g => g.First());

                var dbData = File.ReadAllLines(filePathDB, Encoding.UTF8)
                    .Skip(1)
                    .Select(line => line.Split(','))
                    .GroupBy(parts => int.Parse(parts[1]))
                    .ToDictionary(g => g.Key, g => g.First());

                TrySaveToFile(() =>
                {
                    using (var sw = new StreamWriter(filePathARDeleted, false, Encoding.UTF8))
                    {
                        sw.WriteLine("Project Name,Element ID,Element Type,Element Name,Level,Time,User");
                        foreach (var arEntry in arData)
                        {
                            if (dbData.TryGetValue(arEntry.Key, out var dbEntry))
                            {
                                sw.WriteLine($"{arEntry.Value[0]},{arEntry.Key},{dbEntry[2]},{dbEntry[3]},{dbEntry[4]},{arEntry.Value[2]},{arEntry.Value[3]}");
                            }
                        }
                    }
                }, filePathARDeleted);

                elementsDeleted = false; // Reset the flag after processing
            }, filePathARDeleted);
        }

        private void TrySaveToFile(Action saveAction, string filePath)
        {
            const int maxRetries = 2;
            int retries = 0;
            bool success = false;

            while (!success && retries < maxRetries)
            {
                try
                {
                    saveAction();
                    success = true;
                }
                catch (IOException)
                {
                    retries++;
                    Task.Delay(180000).Wait(); // Wait for 3 minutes before retrying
                }
            }
        }

        private void ExportProjectElements()
        {
            if (currentDocument == null) return;

            string projectName = Path.GetFileNameWithoutExtension(currentDocument.PathName);
            string folderName = $"{projectName}_{userPrefix}";
            string folderPath = Path.Combine(sharedFolderPath, folderName);
            string sourceTablesFolder = Path.Combine(folderPath, "SourceTables");
            Directory.CreateDirectory(folderPath);
            Directory.CreateDirectory(sourceTablesFolder);

            string filePath = Path.Combine(sourceTablesFolder, $"{folderName}_Db.csv");

            var collector = new FilteredElementCollector(currentDocument);
            collector.WhereElementIsNotElementType();

            var elementData = collector.Select(element =>
            {
                string elementType = element.Category?.Name ?? "Unknown";
                string elementName = element.Name;
                string level = element.LevelId != ElementId.InvalidElementId
                    ? currentDocument.GetElement(element.LevelId)?.Name
                    : "Unknown";

                return new Tuple<string, int, string, string, string>(
                    projectName, element.Id.IntegerValue, elementType, elementName, level);
            }).ToList();

            TrySaveToFile(() =>
            {
                using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    sw.WriteLine("Project Name,Element ID,Element Type,Element Name,Level");
                    foreach (var entry in elementData)
                    {
                        sw.WriteLine($"{entry.Item1},{entry.Item2},{entry.Item3},{entry.Item4},{entry.Item5}");
                    }
                }
            }, filePath);
        }
    }
}
