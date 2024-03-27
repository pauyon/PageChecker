using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public class CsvReaderUtility : ReaderBase, IReaderUtility
    {
        public void AnalyzeAndExportResults(string folderPath, string marketClientSheetFilename, string salesSheetFilename)
        {
            throw new NotImplementedException();
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
