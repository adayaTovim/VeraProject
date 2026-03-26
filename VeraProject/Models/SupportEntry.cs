namespace VeraProject.Models;

public class SupportEntry
{
    public string ServiceType { get; set; } = string.Empty;   // Management or Engineer
    public string TicketReference { get; set; } = string.Empty; // Jira/Confluence link
    public string Subject { get; set; } = string.Empty;
    public string HandledBy { get; set; } = string.Empty;
    public string InitiatedBy { get; set; } = string.Empty;
    public string TicketStatus { get; set; } = string.Empty;
    public double Hours { get; set; }
    public DateTime Date { get; set; } = DateTime.Today;
}
