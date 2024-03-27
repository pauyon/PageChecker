using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;

// Display app title
Utility.ShowAppTitle("Page Checker");
Utility.ShowInstructions();

// App instructions
ConsoleUtility.ShowChecklist();

// Check if user is ready
if (!ConsoleUtility.CheckListCompletePrompt()) return;

// Get workspace folder path
IFileReaderUtility fileReaderUtility = new CsvReaderUtility();

var workspaceFolderPath = ConsoleUtility.WorkspaceFolderPrompt().EscapeMarkup();
fileReaderUtility.SetWorkspaceDirectoryPath(workspaceFolderPath);

var workspacePath = Utility.GetWorkspaceFolderPathPrompt();

if (string.IsNullOrEmpty(workspacePath))
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

XmlReaderUtility.SetRootDirectoryPath(workspacePath);

if (XmlReaderUtility.GetRootDirectoryFolders().Count == 0)
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

Utility.WriteSpacedLine($"Workspace path: {XmlReaderUtility.RootDirectory.FullName}");

var foldersToAnalyze = Utility.GetFoldersToAnalyzePrompt();

if (foldersToAnalyze.Count == 0)
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

ConsoleUtility.WriteSpacedLine($"Workspace path: {fileReaderUtility.WorkspaceDirectory.FullName}");

// Get list of folder to run analysis on
var folderNames = ConsoleUtility.SelectWorkspaceFoldersPrompt(fileReaderUtility);

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Processing folders...", ctx =>
    {
        AnsiConsole.WriteLine();

        ConsoleUtility.AnalyzeAndExportResults(fileReaderUtility, folderNames, ".csv");
        ConsoleUtility.WriteSpacedLine("Analysis Complete!");
    });


