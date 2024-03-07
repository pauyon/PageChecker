using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;

AnsiConsole.Write(
    new FigletText("Page Checker")
        .Centered()
        .Color(Color.Green));

Utility.ShowInstructions();

if (!Utility.AreFilesReady())
{
    return;
}

var workspacePath = Utility.GetWorkspaceFolderPath();

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

var foldersToAnalyze = Utility.GetFoldersToAnalyze();

if (foldersToAnalyze.Count == 0)
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Analyzing files...", ctx =>
    {
        AnsiConsole.WriteLine();
        Utility.AnalyzeFolders(foldersToAnalyze);
        Utility.WriteSpacedLine("Analysis Complete!");
    });


