using System.Drawing;

namespace PageChecker.Domain.Models;

public class MarketClient
{
    public bool PassedCheck { get; set; } = false;
    public string CustomerName { get; set; }
    public double Size { get; set; }
    public string Rep { get; set; }
    public string Categories { get; set; }
    public string ContractStatus { get; set; }
    public string Artwork { get; set; }
    public string Notes { get; set; }
    public string Placement { get; set; }
    public string AccountingCustomerName { get; set; }

    public string AccurateCustomerName
    {
        get
        {
            if (!string.IsNullOrEmpty(AccountingCustomerName) && AccountingCustomerName.ToLower() != CustomerName.ToLower())
            {
                return AccountingCustomerName;
            }
            else
            {
                return CustomerName;
            }
        }
    }
}
