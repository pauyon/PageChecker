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

var xmlReaderUtility = new XmlReaderUtility();

xmlReaderUtility.SetRootDirectoryPath(Utility.GetWorkspaceFolderPath().EscapeMarkup());

if (xmlReaderUtility.GetRootDirectoryFolders().Count() == 0)
{
    Utility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

Utility.WriteSpacedLine($"Workspace path: {xmlReaderUtility.RootDirectory.FullName}");

var foldersToAnalyze = Utility.GetFoldersToAnalyze(xmlReaderUtility);

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
        Utility.AnalyzeFolders(xmlReaderUtility, foldersToAnalyze);
        Utility.WriteSpacedLine("Analysis Complete!");
    });


