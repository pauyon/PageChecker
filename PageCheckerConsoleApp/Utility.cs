using PageChecker.Library;
using Spectre.Console;

namespace PageChecker.ConsoleApp;

public static class Utility
{
    public static void ShowInstructions()
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("Before starting, ensure the following:");
        AnsiConsole.MarkupLine("======================================");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("1. Create a [green]workspace[/] folder. This will contain all [green]client[/] folders for analysis.");
        AnsiConsole.MarkupLine("2. Within each [green]client[/] folder include the [green]market spreadsheet[/] and [green]sales run spreadsheet[/] files.");
        AnsiConsole.MarkupLine("3. Ensure the spreadsheets in each folder are in excel format ([green].xlsx[/]).");
        AnsiConsole.MarkupLine("4. Ensure the sales spreadsheets contains the keywords [green]'sales sheet'[/] OR [green]'salessheet'[/].");
        AnsiConsole.MarkupLine("5. Ensure there is only [green]1 market/client spreadsheet[/] and [green]1 sales run spreadsheet[/] per folder.");
        AnsiConsole.WriteLine("");
    }

    public static bool AreFilesReady()
    {
        if (!AnsiConsole.Confirm("Have you completed everything on the checklist?"))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("The application will now close. Run again once the files are setup.");
            return false;
        }

        return true;
    }

    public static string GetWorkspaceRootPath()
    {
        string rootPath = "./";

        if (!AnsiConsole.Confirm("Is the workspace folder located in the same place as the app?"))
        {
            rootPath = AnsiConsole.Ask<string>("Path of workspace folder:");
        }

        return rootPath;
    }

    public static List<string> GetWorkspaceRootPathFolders(string path)
    {
        var directory = new DirectoryInfo(path);

        var folders = directory.GetDirectories().Select(x => x.Name).ToList();
        return folders;
    }

    public static string GetWorkspaceFolderPath()
    {
        var rootPath = GetWorkspaceRootPath();
        var folders = GetWorkspaceRootPathFolders(rootPath);

        var folderName = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
            .Title("Which folder is the [green]workspace[/] folder?")
            .PageSize(10)
            .MoreChoicesText("[grey](Move up and down to reveal more files)[/]")
            .AddChoices(folders));

        return Path.Combine(rootPath, folderName);
    }

    public static List<string> GetFoldersToAnalyze(XmlReaderUtility xmlReaderUtility)
    {
        List<string> foldersToAnalyze = xmlReaderUtility.GetRootDirectoryFolders();

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
                .AddChoices(xmlReaderUtility.GetRootDirectoryFolders()));
        }

        return foldersToAnalyze;
    }

    public static List<string> GetFolderFiles(string path)
    {
        var directory = new DirectoryInfo(path);

        var files = directory.GetFiles().Where(x => !x.Name.ToLower().Contains("results")).Select(x => x.Name).ToList();

        return files;
    }

    public static void AnalyzeFolders(XmlReaderUtility xmlReaderUtility, List<string> folders)
    {
        foreach (string folder in folders)
        {
            var folderPath = Path.Combine(xmlReaderUtility.RootDirectory.FullName, folder);
            var files = GetFolderFiles(folderPath);

            if (files == null || files.Count == 0)
            {
                AnsiConsole.MarkupLine($"Folder [green]{folder}[/] skipped. No spreadsheets found.");
                continue;
            }

            if (files.Count > 2)
            {
                AnsiConsole.MarkupLine($"Folder [green]{folder}[/] skipped. Too many spreadsheets found.");
                continue;
            }

            var salesSheet = files.First(x => x.ToLower().Contains("salessheet") || x.ToLower().Contains("sales sheet"));
            var marketSheet = files.First(x => !x.ToLower().Contains("salessheet") || !x.ToLower().Contains("sales sheet"));

            xmlReaderUtility.OpenSalesSheet(Path.Combine(folderPath, salesSheet));
            xmlReaderUtility.OpenMarketSheet(Path.Combine(folderPath, marketSheet));

            xmlReaderUtility.ExportResults(folderPath);
            AnsiConsole.MarkupLine($"Folder [green]{folder}[/] analyzed successfully.");
        }
    }

    public static void WriteSpacedLine(string text)
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(text);
        AnsiConsole.WriteLine();
    }
}
