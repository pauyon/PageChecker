using PageChecker.Library;
using Spectre.Console;

namespace PageCheckerConsoleApp;

public static class Utility
{
    public static void ShowInstructions()
    {
        AnsiConsole.WriteLine("");
        AnsiConsole.MarkupLine("Before starting, ensure the following:");
        AnsiConsole.MarkupLine("===================================");
        AnsiConsole.MarkupLine("1. Create a [green]workspace folder[/] that will contain [green]client folders[/] to be analyzed.");
        AnsiConsole.MarkupLine("2. Inside each [green]client folder[/] include the [green]market spreadsheet[/] and [green]sales run[/].");
        AnsiConsole.MarkupLine("3. Ensure the spreadsheets are in excel ([green].xlsx[/]) format.");
        AnsiConsole.MarkupLine("4. Ensure the sales spreadsheets contains the keywords [green]'sales sheet'[/] or [green]'salessheet'[/].");
        AnsiConsole.WriteLine("");
    }

    public static bool AreFilesReady()
    {
        if (!AnsiConsole.Confirm("Have you prepped the folders and files?"))
        {
            AnsiConsole.MarkupLine("");
            AnsiConsole.MarkupLine("The app will now close, run again once the files are setup.");
            return false;
        }

        return true;
    }

    public static string GetWorkspaceRootPath()
    {
        string rootPath = "./";

        if (!AnsiConsole.Confirm("Is the workspace folder in the same folder as the app?"))
        {
            rootPath = AnsiConsole.Ask<string>("Path of workspace folder :");
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

    public static List<string> GetFoldersToAnalyze()
    {
        if (!AnsiConsole.Confirm("Would you like to run the analyzer on all folders in the workspace?"))
        {
            List<string> folders = AnsiConsole.Prompt(
            new MultiSelectionPrompt<string>()
                .Title("Which folders would you like to run the analyzer on?")
                .Required()
                .PageSize(10)
                .MoreChoicesText("[grey](Move up and down to reveal more folders)[/]")
                .InstructionsText(
                    "[grey](Press [blue]<space>[/] to toggle a folder, " +
                    "[green]<enter>[/] to accept)[/]")
                .AddChoices(XmlReaderUtility.GetRootDirectoryFolders()));
            return folders;
        }

        return XmlReaderUtility.GetRootDirectoryFolders();
    }

    public static List<string> GetFolderFiles(string path)
    {
        var directory = new DirectoryInfo(path);

        var files = directory.GetFiles().Where(x => !x.Name.ToLower().Contains("results")).Select(x => x.Name).ToList();

        return files;
    }

    public static void AnalyzeFolders(List<string> folders)
    {
        foreach (string folder in folders)
        {
            var folderPath = Path.Combine(XmlReaderUtility.RootDirectory.FullName, folder);
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

            XmlReaderUtility.OpenSalesSheet(Path.Combine(folderPath, salesSheet));
            XmlReaderUtility.OpenMarketSheet(Path.Combine(folderPath, marketSheet));

            XmlReaderUtility.ExportResults(folderPath);
            AnsiConsole.MarkupLine($"Folder [green]{folder}[/] analyzed successfully.");
        }
    }
}
