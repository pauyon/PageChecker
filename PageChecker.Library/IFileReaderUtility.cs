using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IFileReaderUtility : IFileReaderBase
    {
        List<MarketClient> GetMarketClientSheetData();
        List<SalesRun> GetSalesSheetData();
        void AnalyzeAndExportResults(string folderPath, string marketClientSheetFilename, string salesSheetFilename);
    }
}
