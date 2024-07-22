using Microsoft.Extensions.Logging;
using PageChecker.Library;
using Spectre.Console;

namespace PageChecker.ConsoleApp;

public class App
{
    private readonly ILogger _logger;

    public App(ILogger<App> logger)
    {
        _logger = logger;
    }

    public void Run()
    {
        _logger.LogInformation("Application started.");
        var consoleUtility = new ConsoleUtility(_logger);

        // Display app title
        consoleUtility.ShowAppTitle("Page Checker");

        // App instructions
        consoleUtility.ShowChecklist();

        // Check if user is ready
        if (!consoleUtility.CheckListCompletePrompt())
        {
            _logger.LogInformation("User has not completed checklist. Closing app.");
            return;
        }

        // Get workspace folder path
        IFileReaderUtility fileReaderUtility = new CsvReaderUtility(_logger);

        var workspaceFolderPath = consoleUtility.WorkspaceFolderPrompt().EscapeMarkup();
        fileReaderUtility.SetWorkspaceDirectoryPath(workspaceFolderPath);

        // Check workspace folder structure
        if (!fileReaderUtility.GetWorkspaceFolders().Any())
        {
            _logger.LogInformation($"There were no folders in workspace. Closing application.");
            consoleUtility.WriteSpacedLine($"There were no folders in workspace. Closing application.");
            return;
        }

        consoleUtility.WriteSpacedLine($"Workspace path: {fileReaderUtility.WorkspaceDirectory.FullName}");

        // Get list of folder to run analysis on
        var folderNames = consoleUtility.SelectWorkspaceFoldersPrompt(fileReaderUtility);

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Star)
            .SpinnerStyle(Style.Parse("green bold"))
            .Start("Analyzing files...", ctx =>
            {
                AnsiConsole.WriteLine();
                consoleUtility.AnalyzeAndExportResults(fileReaderUtility, folderNames, ".csv");
                consoleUtility.WriteSpacedLine("Analysis Complete!");
            });

        _logger.LogInformation("Application finished.");
    }
}
