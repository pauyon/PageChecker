using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IFileReaderBase
    {
        DirectoryInfo WorkspaceDirectory { get; set; }
        List<string> MarketClientSheetHeaders { get; set; }
        List<string> SalesSheetHeaders { get; set; }

        double GetPageSizeNumericValue(string pageDescription);
        void SetWorkspaceDirectoryPath(string directoryPath);
        List<string> GetWorkspaceFolders();
        void RemoveExistingResultsExcel(string resultsSheetFilePath);
        List<MarketClient> CompareSheetsData(List<MarketClient> marketClientSheetData, List<SalesRun> salesRunSheetData);
        void GenerateResultsExcel(List<MarketClient> checkedMarketData, string resultsExportPath);
        List<string> RenameAccountingColumn(List<string> headers);
    }
}