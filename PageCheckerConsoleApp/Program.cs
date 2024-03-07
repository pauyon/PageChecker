using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;


Utility.ShowAppTitle("Page Checker");
Utility.ShowInstructions();

if (!Utility.FilesReadyPrompt())
{
    return;
}

var workspacePath = Utility.GetWorkspaceFolderPathPrompt();

if (string.IsNullOrEmpty(workspacePath))
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

XmlReaderUtility.SetRootDirectoryPath(workspacePath);

if (XmlReaderUtility.GetRootDirectoryFolders().Count() == 0)
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

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Processing folders...", ctx =>
    {
        AnsiConsole.WriteLine();
        Utility.ProcessFolders(foldersToAnalyze);
        Utility.WriteSpacedLine("Analysis Complete!");
    });


