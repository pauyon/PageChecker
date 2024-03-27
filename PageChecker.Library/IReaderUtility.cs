using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IReaderUtility : IReaderBase
    {
        List<Market> GetMarketSheetData();
        List<SalesRun> GetSalesSheetData();
        void AnalyzeAndExportResults(string folderPath, string marketSheetFilename, string salesSheetFilename);
    }
}
