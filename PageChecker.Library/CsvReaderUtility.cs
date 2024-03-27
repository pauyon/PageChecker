using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public class CsvReaderUtility : ReaderBase, IReaderUtility
    {
        public void AnalyzeAndExportResults(string folderPath, string marketClientSheetFilename, string salesSheetFilename)
        {
            var resultsExportPath = Path.Combine(folderPath, "Results.xlsx");

            RemoveExistingResultsExcel(resultsExportPath);

            var marketClientSheetData = GetMarketClientSheetData();
            var salesRunSheetData = GetSalesSheetData();
            var checkedMarketData = CompareSheetsData(marketClientSheetData, salesRunSheetData);

            GenerateResultsExcel(checkedMarketData, resultsExportPath);
        }

        public List<MarketClient> GetMarketClientSheetData()
        {
            var marketData = new List<MarketClient>();

            SalesSheetHeaders = new();

            return marketData;
        }

        public List<SalesRun> GetSalesSheetData()
        {
            var salesRunData = new List<SalesRun>();

            SalesSheetHeaders = new();

            return salesRunData;
        }
    }
}
