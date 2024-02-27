using System.Drawing;

namespace PageChecker.Domain.Models;

public class Market
{
    public bool PassedCheck { get; set; } = false;
    public string Customer { get; set; }
    public double Size { get; set; }
    public string Rep { get; set; }
    public string Categories { get; set; }
    public string ContractStatus { get; set; }
    public string Artwork { get; set; }
    public string Notes { get; set; }
    public string Placement { get; set; }
    public string AccountingNotes { get; set; }
}
