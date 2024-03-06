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

XmlReaderUtility.SetRootDirectoryPath(Utility.GetWorkspaceFolderPath().EscapeMarkup());

AnsiConsole.WriteLine();
AnsiConsole.MarkupLine($"Workspace path set to: {XmlReaderUtility.RootDirectory.FullName}");
AnsiConsole.WriteLine();

var foldersToAnalyze = Utility.GetFoldersToAnalyze();

if (foldersToAnalyze.Count == 0)
{
    AnsiConsole.WriteLine();
    AnsiConsole.MarkupLine($"There were no folders in workspace. App closing.");
    AnsiConsole.WriteLine();

    return;
}

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Analyzing files...", ctx =>
    {
        Utility.AnalyzeFolders(foldersToAnalyze);
        
        Console.WriteLine("Analysis complete!");
    });


