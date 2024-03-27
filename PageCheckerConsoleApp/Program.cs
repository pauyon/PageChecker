using PageChecker.Library;
using PageChecker.ConsoleApp;
using Spectre.Console;

AnsiConsole.Write(
    new FigletText("Page Checker")
        .Centered()
        .Color(Color.Green));

ConsoleUtility.ShowInstructions();

if (!ConsoleUtility.AreFilesReady())
{
    return;
}

var xmlReaderUtility = new XmlReaderUtility();

xmlReaderUtility.SetRootDirectoryPath(ConsoleUtility.GetWorkspaceFolderPath().EscapeMarkup());

if (!xmlReaderUtility.GetRootDirectoryFolders().Any())
{
    ConsoleUtility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

ConsoleUtility.WriteSpacedLine($"Workspace path: {xmlReaderUtility.RootDirectory.FullName}");

var foldersToAnalyze = ConsoleUtility.GetFoldersToAnalyze(xmlReaderUtility);

if (foldersToAnalyze.Count == 0)
{
    ConsoleUtility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
    return;
}

AnsiConsole.Status()
    .Spinner(Spinner.Known.Star)
    .SpinnerStyle(Style.Parse("green bold"))
    .Start("Analyzing files...", ctx =>
    {
        AnsiConsole.WriteLine();
        ConsoleUtility.AnalyzeAndExportResults(xmlReaderUtility, foldersToAnalyze);
        ConsoleUtility.WriteSpacedLine("Analysis Complete!");
    });


