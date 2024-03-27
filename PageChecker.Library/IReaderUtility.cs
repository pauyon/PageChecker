using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IReaderUtility : IReaderBase
    {
        List<MarketClient> GetMarketClientSheetData();
        List<SalesRun> GetSalesSheetData();
        void AnalyzeAndExportResults(string folderPath, string marketClientSheetFilename, string salesSheetFilename);
    }
}
