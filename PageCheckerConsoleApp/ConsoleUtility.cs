using PageChecker.Library;
using Spectre.Console;

namespace PageChecker.ConsoleApp;

public static class ConsoleUtility
{
    /// <summary>
    /// Write instructions for folder structure to console.
    /// </summary>
    public static void ShowChecklist()
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("Before starting, ensure the following:");
        AnsiConsole.MarkupLine("======================================");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("1. Create a [green]workspace[/] folder. This will contain all [green]client[/] folders for analysis.");
        AnsiConsole.MarkupLine("2. Within each [green]client[/] folder include the [green]market spreadsheet[/] and [green]sales run spreadsheet[/] files.");
        AnsiConsole.MarkupLine("3. Ensure the spreadsheets in each folder are in excel format ([green].xlsx[/]).");
        AnsiConsole.MarkupLine("4. Ensure the sales spreadsheets contains the keywords [green]'sales run sheet'[/] OR [green]'salesrunsheet'[/].");
        AnsiConsole.MarkupLine("5. Ensure there is only [green]1 market/client spreadsheet[/] and [green]1 sales run spreadsheet[/] per folder.");
        AnsiConsole.WriteLine("");
    }

    /// <summary>
    /// Prompts user if the checklist has been complete.
    /// </summary>
    /// <returns>Bool value of true or false.</returns>
    public static bool CheckListCompletePrompt()
    {
        if (!AnsiConsole.Confirm("Have you completed everything on the checklist?"))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("The application will now close. Run again once the files are setup.");
            return false;
        }

        return true;
    }

    /// <summary>
    /// Prompts user for a workspace folder path. Returns working directory path by default.
    /// </summary>
    /// <returns>Full path of workspace folder.</returns>
    public static string WorkspacePathPrompt()
    {
        string rootPath = "./";

        if (!AnsiConsole.Confirm("Is the workspace folder located in the same place as the app?"))
        {
            rootPath = AnsiConsole.Ask<string>("Path of workspace folder:");
        }

        return rootPath;
    }

    /// <summary>
    /// Gets a list of folder names located within the workspace folder.
    /// </summary>
    /// <param name="path">Path of workspace folder.</param>
    /// <returns>List of folder names of workspace folder path.</returns>
    public static List<string> GetWorkspaceRootPathFolders(string path)
    {
        var directory = new DirectoryInfo(path);

        var folders = directory.GetDirectories().Select(x => x.Name).ToList();
        return folders;
    }

    /// <summary>
    /// Prompts user for the path of the workspace folder to work with.
    /// </summary>
    /// <returns>Full path of workspace folder.</returns>
    public static string WorkspaceFolderPrompt()
    {
        var rootPath = WorkspacePathPrompt();
        var folders = GetWorkspaceRootPathFolders(rootPath);

        var folderName = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
            .Title("Which folder is the [green]workspace[/] folder?")
            .PageSize(10)
            .MoreChoicesText("[grey](Move up and down to reveal more files)[/]")
            .AddChoices(folders));

        return Path.Combine(rootPath, folderName);
    }

    /// <summary>
    /// Prompts user for list of folders within workspace folder to analyze.
    /// All folders selected by default.
    /// </summary>
    /// <param name="xmlReaderUtility">Tool for reading spreadsheets.</param>
    /// <returns>List of folder names to analyze.</returns>
    public static List<string> SelectWorkspaceFoldersPrompt(XmlReaderUtility xmlReaderUtility)
    {
        List<string> foldersToAnalyze = xmlReaderUtility.GetWorkspaceFolders();

        if (!AnsiConsole.Confirm($"[green]{foldersToAnalyze.Count}[/] folders were found in workspace. Run analyzer on all files?"))
        {
            foldersToAnalyze = AnsiConsole.Prompt(
            new MultiSelectionPrompt<string>()
                .Title($"{Environment.NewLine}Which folders would you like to run the analyzer on?")
                .Required()
                .PageSize(10)
                .MoreChoicesText("[grey](Move up and down to reveal more folders)[/]")
                .InstructionsText(
                    "[grey](Press [blue]<space>[/] to toggle a folder, " +
                    "[green]<enter>[/] to accept)[/]")
                .AddChoices(xmlReaderUtility.GetWorkspaceFolders()));
        }

        return foldersToAnalyze;
    }

    /// <summary>
    /// Return list of files within a given folder path and extension.
    /// </summary>
    /// <param name="path">Folder path.</param>
    /// <param name="validExtension">File extension.</param>
    /// <returns></returns>
    public static List<string> GetFolderFiles(string path, string validExtension)
    {
        var directory = new DirectoryInfo(path);
        var files = directory.GetFiles().Where(x => !x.Name.ToLower().Contains("results") && x.Extension == validExtension).Select(x => x.Name).ToList();
        return files;
    }

    /// <summary>
    /// Analyze folders selected and export results to excel file.
    /// </summary>
    /// <param name="xmlReaderUtility">Tool for reading spreadsheet data.</param>
    /// <param name="folderNames">List of folders to analyze.</param>
    public static void AnalyzeAndExportResults(IReaderUtility xmlReaderUtility, List<string> folderNames)
    {
        foreach (string folderName in folderNames)
        {
            var fullFolderPath = Path.Combine(xmlReaderUtility.WorkspaceDirectory.FullName, folderName);
            var files = GetFolderFiles(fullFolderPath, ".xlsx");

            if (files == null || files.Count == 0)
            {
                AnsiConsole.MarkupLine($"Folder [green]{folderName}[/] skipped. No spreadsheets found.");
                continue;
            }

            if (files.Count > 2)
            {
                AnsiConsole.MarkupLine($"Folder [green]{folderName}[/] skipped. Too many spreadsheets found.");
                continue;
            }

            var salesRunSheetFilename = files.FirstOrDefault(x => x.ToLower().Replace(" ", "").Contains("salesrunsheet"));

            if (salesRunSheetFilename == null)
            {
                AnsiConsole.MarkupLine($"Folder [green]{folderName}[/] skipped. Salesrun sheet is missing.");
                continue;
            }
            else
            {
                files.Remove(salesRunSheetFilename);
            }

            var marketSheetFilename = files[0];

            xmlReaderUtility.ExportResults(fullFolderPath, marketSheetFilename, salesRunSheetFilename);
            AnsiConsole.MarkupLine($"Folder [green]{folderName}[/] analyzed successfully.");
        }
    }

    /// <summary>
    /// Write text to console with a new line above and below text.
    /// </summary>
    /// <param name="text"></param>
    public static void WriteSpacedLine(string text)
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(text);
        AnsiConsole.WriteLine();
    }
}
