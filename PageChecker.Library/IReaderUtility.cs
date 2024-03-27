using ClosedXML.Excel;
using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public interface IReaderUtility
    {
        public List<Market> GetMarketSheetData();
        List<SalesRun> GetSalesSheetData();
        List<Market> CompareSheetsData(List<Market> marketSheetData, List<SalesRun> salesSheetData);
        void OpenSalesSheet(string filename);
        void OpenMarketSheet(string filename);
        void ExportResults(string folderPath);
    }
}
