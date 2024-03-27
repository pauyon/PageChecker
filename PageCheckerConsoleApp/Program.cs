using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;

// Display app title
ConsoleUtility.ShowAppTitle("Page Checker");

// App instructions
ConsoleUtility.ShowChecklist();

// Check if user is ready
if (!ConsoleUtility.CheckListCompletePrompt()) return;

// Get workspace folder path
IFileReaderUtility fileReaderUtility = new CsvReaderUtility();

var workspaceFolderPath = ConsoleUtility.WorkspaceFolderPrompt().EscapeMarkup();
fileReaderUtility.SetWorkspaceDirectoryPath(workspaceFolderPath);

// Check workspace folder structure
if (!fileReaderUtility.GetWorkspaceFolders().Any())
{
    ConsoleUtility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

ConsoleUtility.WriteSpacedLine($"Workspace path: {fileReaderUtility.WorkspaceDirectory.FullName}");

// Get list of folder to run analysis on
var folderNames = ConsoleUtility.SelectWorkspaceFoldersPrompt(fileReaderUtility);

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Analyzing files...", ctx =>
    {
        AnsiConsole.WriteLine();
        ConsoleUtility.AnalyzeAndExportResults(fileReaderUtility, folderNames, ".csv");
        ConsoleUtility.WriteSpacedLine("Analysis Complete!");
    });


