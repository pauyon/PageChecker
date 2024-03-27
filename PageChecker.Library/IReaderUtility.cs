using ClosedXML.Excel;
using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IReaderUtility
    {
        List<Market> GetMarketSheetData();
        List<SalesRun> GetSalesSheetData();
        void ExportResults(string folderPath, string marketSheetFilename, string salesSheetFilename);
    }
}
