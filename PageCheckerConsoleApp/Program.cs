using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;

// Display app title
AnsiConsole.Write(
    new FigletText("Page Checker")
        .Centered()
        .Color(Color.Green));

// App instructions
ConsoleUtility.ShowChecklist();

// Check if user is ready
if (!ConsoleUtility.CheckListCompletePrompt()) return;

// Get workspace folder path
IReaderUtility xmlReaderUtility = new XmlReaderUtility();

var workspaceFolderPath = ConsoleUtility.WorkspaceFolderPrompt().EscapeMarkup();
xmlReaderUtility.SetWorkspaceDirectoryPath(workspaceFolderPath);

// Check workspace folder structure
if (!xmlReaderUtility.GetWorkspaceFolders().Any())
{
    ConsoleUtility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

ConsoleUtility.WriteSpacedLine($"Workspace path: {xmlReaderUtility.WorkspaceDirectory.FullName}");

// Get list of folder to run analysis on
var folderNames = ConsoleUtility.SelectWorkspaceFoldersPrompt(xmlReaderUtility);

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Analyzing files...", ctx =>
    {
        AnsiConsole.WriteLine();
        ConsoleUtility.AnalyzeAndExportResults(xmlReaderUtility, folderNames);
        ConsoleUtility.WriteSpacedLine("Analysis Complete!");
    });


