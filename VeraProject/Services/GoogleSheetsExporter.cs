using System.Net.Http;
using System.Text;
using System.Text.Json;
using VeraProject.Models;

namespace VeraProject.Services;

public static class GoogleSheetsExporter
{
    private static readonly string BasePath = AppDomain.CurrentDomain.BaseDirectory;
    private static readonly string ConfigFile = Path.Combine(BasePath, "google-sheet-config.txt");

    private static readonly HttpClient Http = new();

    // Config file has 2 lines:
    // Line 1: Apps Script web app URL
    // Line 2: Google Sheet URL (for opening in browser)
    public static bool IsConfigured => File.Exists(ConfigFile) && GetConfig() != null;

    public static string GetSheetUrl()
    {
        var config = GetConfig();
        return config?.sheetUrl ?? "";
    }

    public static void AppendEntry(SupportEntry entry)
    {
        var config = GetConfig()
            ?? throw new InvalidOperationException("Google Sheet not configured.");

        var payload = JsonSerializer.Serialize(new
        {
            entry.ServiceType,
            entry.TicketReference,
            entry.Subject,
            entry.HandledBy,
            entry.InitiatedBy,
            entry.TicketStatus,
            entry.Hours,
            Date = entry.Date.ToString("yyyy-MM-dd")
        });

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        var response = Http.PostAsync(config.webAppUrl, content).GetAwaiter().GetResult();

        if (!response.IsSuccessStatusCode)
            throw new Exception($"Google Sheets returned {response.StatusCode}");
    }

    private static (string webAppUrl, string sheetUrl)? GetConfig()
    {
        if (!File.Exists(ConfigFile)) return null;

        var lines = File.ReadAllLines(ConfigFile)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToArray();

        if (lines.Length < 2) return null;
        return (lines[0].Trim(), lines[1].Trim());
    }
}
