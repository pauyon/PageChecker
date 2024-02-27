namespace PageChecker.Domain.Models;
public class SalesRun
{
    public string Client { get; set; }
    public string Product { get; set; }
    public string Description { get; set; }
    public string SalesRep { get; set; }
    public decimal Net { get; set; }
    public int Barter { get; set; }
}
