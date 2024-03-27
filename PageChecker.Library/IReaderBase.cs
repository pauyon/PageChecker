using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IReaderBase
    {
        public DirectoryInfo WorkspaceDirectory { get; set; }
        public List<string> MarketSheetHeaders { get; set; }
        public List<string> SalesSheetHeaders { get; set; }

        public double GetPageSizeNumericValue(string pageDescription);
        public void SetWorkspaceDirectoryPath(string directoryPath);
        public List<string> GetWorkspaceFolders();
        public void RemoveExistingResultsExcel(string resultsSheetFilePath);
        public List<Market> CompareSheetsData(List<Market> marketSheetData, List<SalesRun> salesRunSheetData);
        public void GenerateResultsExcel(List<Market> checkedMarketData, string resultsExportPath);
    }
}