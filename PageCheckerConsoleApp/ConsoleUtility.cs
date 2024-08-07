﻿using Microsoft.Extensions.Logging;
using PageChecker.Library;
using Spectre.Console;

namespace PageChecker.ConsoleApp;

public class ConsoleUtility
{
    private readonly ILogger _logger;

    public ConsoleUtility(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Write instructions for folder structure to console.
    /// </summary>
    public void ShowChecklist()
    {
        _logger.LogInformation("Displaying checklist");

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("Before starting, ensure the following:");
        AnsiConsole.MarkupLine("======================================");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("1. Create a [green]workspace[/] folder. This will contain all [green]market[/] folders for analysis.");
        AnsiConsole.MarkupLine("2. Within each [green]market[/] folder include the [green]market client spreadsheet[/] and [green]sales run spreadsheet[/] files.");
        AnsiConsole.MarkupLine("3. Ensure the spreadsheets in each folder are in comma separated format ([green].csv[/]).");
        AnsiConsole.MarkupLine("4. Ensure the sales run contains the following keywords: '[green]sales run sheet[/]', '[green]salesrun sheet[/]', or '[green]salesrunsheet'[/]");
        AnsiConsole.MarkupLine("5. Ensure there is only [green]1 market client spreadsheet[/] and [green]1 sales run spreadsheet[/] per folder.");
        AnsiConsole.WriteLine("");
    }

    /// <summary>
    /// Prompts user if the checklist has been complete.
    /// </summary>
    /// <returns>Bool value of true or false.</returns>
    public bool CheckListCompletePrompt()
    {
        if (!AnsiConsole.Confirm("Have you completed everything on the checklist?"))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("The application will now close. Run again once the files are setup.");
            _logger.LogInformation("User has NOT completed checklist completion. Closing app");
            return false;
        }

        _logger.LogInformation("User has completed checklist.");
        return true;
    }

    /// <summary>
    /// Prompts user for a workspace folder path. Returns working directory path by default.
    /// </summary>
    /// <returns>Full path of workspace folder.</returns>
    public string WorkspacePathPrompt()
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
    public List<string> GetWorkspaceRootPathFolders(string path)
    {
        var directory = new DirectoryInfo(path);

        var folders = directory.GetDirectories().Select(x => x.Name).ToList();
        return folders;
    }

    /// <summary>
    /// Prompts user for the path of the workspace folder to work with.
    /// </summary>
    /// <returns>Full path of workspace folder.</returns>
    public string WorkspaceFolderPrompt()
    {
        var rootPath = WorkspacePathPrompt();
        var folders = GetWorkspaceRootPathFolders(rootPath);

        var folderName = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
            .Title("Which folder is the [green]workspace[/] folder?")
            .PageSize(10)
            .MoreChoicesText("[grey](Move up and down to reveal more files)[/]")
            .AddChoices(folders));

        var workspacePath = Path.Combine(rootPath, folderName);
        _logger.LogInformation($"Workspace path set to : {workspacePath}");

        return workspacePath;
    }

    /// <summary>
    /// Prompts user for list of folders within workspace folder to analyze.
    /// All folders selected by default.
    /// </summary>
    /// <param name="xmlReaderUtility">Tool for reading spreadsheets.</param>
    /// <returns>List of folder names to analyze.</returns>
    public List<string> SelectWorkspaceFoldersPrompt(IFileReaderUtility xmlReaderUtility)
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

        _logger.LogInformation("Selected folders: " + string.Join(',', foldersToAnalyze));

        return foldersToAnalyze;
    }

    /// <summary>
    /// Return list of files within a given folder path and extension.
    /// </summary>
    /// <param name="path">Folder path.</param>
    /// <param name="validExtension">File extension.</param>
    /// <returns></returns>
    public List<string> GetFolderFiles(string path, string validExtension)
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
    /// <param name="fileExtension">Extension of files to analyze.</param>
    public void AnalyzeAndExportResults(IFileReaderUtility xmlReaderUtility, List<string> folderNames, string fileExtension)
    {
        foreach (string folderName in folderNames)
        {
            var fullFolderPath = Path.Combine(xmlReaderUtility.WorkspaceDirectory.FullName, folderName);
            var files = GetFolderFiles(fullFolderPath, fileExtension);

            if (files == null || files.Count == 0)
            {
                _logger.LogWarning($"Market folder {folderName} skipped. No spreadsheets found.");
                AnsiConsole.MarkupLine($"Market folder [green]{folderName}[/] skipped. No spreadsheets found.");
                continue;
            }

            if (files.Count > 2)
            {
                _logger.LogWarning($"Market folder {folderName} skipped. Too many spreadsheets found.");
                AnsiConsole.MarkupLine($"Market folder [green]{folderName}[/] skipped. Too many spreadsheets found.");
                continue;
            }

            var salesRunSheetFilename = files.FirstOrDefault(x => x.ToLower().Replace(" ", "").Contains("salesrunsheet"));

            if (salesRunSheetFilename == null)
            {
                _logger.LogWarning($"Market folder {folderName} skipped. Salesrun sheet is missing.");
                AnsiConsole.MarkupLine($"Market folder [green]{folderName}[/] skipped. Salesrun sheet is missing.");
                continue;
            }
            else
            {
                files.Remove(salesRunSheetFilename);
            }

            var marketClientSheetFilename = files[0];

            xmlReaderUtility.AnalyzeAndExportResults(
                fullFolderPath, 
                Path.Combine(fullFolderPath, marketClientSheetFilename), 
                Path.Combine(fullFolderPath, salesRunSheetFilename));

            _logger.LogInformation($"Folder {folderName} analyzed successfully.");
            AnsiConsole.MarkupLine($"Folder [green]{folderName}[/] analyzed successfully.");
        }
    }

    /// <summary>
    /// Write text to console with a new line above and below text.
    /// </summary>
    /// <param name="text"></param>
    public void WriteSpacedLine(string text)
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(text);
        AnsiConsole.WriteLine();
    }

    /// <summary>
    /// Write title of app to console.
    /// </summary>
    /// <param name="title"></param>
    public void ShowAppTitle(string title)
    {
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
        string version = "v." + fvi.FileVersion ?? string.Empty;

        _logger.LogInformation("Showing title");
        AnsiConsole.Write(
            new FigletText("------------")
                .Centered()
                .Color(Color.Green));
        AnsiConsole.Write(
            new FigletText(title)
                .Centered()
                .Color(Color.Green));
        AnsiConsole.Write(
            new FigletText(version)
                .Centered()
                .Color(Color.Green));
        AnsiConsole.Write(
            new FigletText("------------")
                .Centered()
                .Color(Color.Green));
    }
}
