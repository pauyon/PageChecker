using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public class CsvReaderUtility : ReaderBase, IReaderUtility
    {
        public List<Market> CompareSheetsData(List<Market> marketSheetData, List<SalesRun> salesSheetData)
        {
            throw new NotImplementedException();
        }

        public void ExportResults(string folderPath)
        {
            throw new NotImplementedException();
        }

        public List<Market> GetMarketSheetData()
        {
            throw new NotImplementedException();
        }

        public List<SalesRun> GetSalesSheetData()
        {
            throw new NotImplementedException();
        }

        public void OpenMarketSheet(string filename)
        {
            throw new NotImplementedException();
        }

        public void OpenSalesSheet(string filename)
        {
            throw new NotImplementedException();
        }
    }
}
