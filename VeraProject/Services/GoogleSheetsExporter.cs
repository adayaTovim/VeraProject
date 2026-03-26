using System.Globalization;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using VeraProject.Models;
using DrivePermission = Google.Apis.Drive.v3.Data.Permission;

namespace VeraProject.Services;

public static class GoogleSheetsExporter
{
    private const string ApplicationName = "VeraProject";
    private const string SpreadsheetTitle = "IUCC-HOURS REPORT";

    private static readonly string BasePath = AppDomain.CurrentDomain.BaseDirectory;
    private static readonly string CredentialPath = Path.Combine(BasePath, "service-account.json");
    private static readonly string SpreadsheetIdFile = Path.Combine(BasePath, "spreadsheet-id.txt");
    private static readonly string ShareEmailFile = Path.Combine(BasePath, "google-share-email.txt");

    private static string? _spreadsheetId;

    public static bool IsConfigured =>
        File.Exists(CredentialPath) && File.Exists(ShareEmailFile);

    public static void AppendEntry(SupportEntry entry)
    {
        var service = CreateSheetsService();
        var spreadsheetId = EnsureSpreadsheetExists(service);

        // Append to Raw Data
        var rawRow = new List<object>
        {
            entry.ServiceType,
            entry.TicketReference,
            entry.Subject,
            entry.HandledBy,
            entry.InitiatedBy,
            entry.TicketStatus,
            entry.Hours,
            entry.Date.ToString("yyyy-MM-dd")
        };

        var appendRequest = service.Spreadsheets.Values.Append(
            new ValueRange { Values = new List<IList<object>> { rawRow } },
            spreadsheetId,
            "'Raw Data'!A:H");
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        appendRequest.Execute();

        // Read all raw data and rebuild sheets
        var entries = ReadAllRawData(service, spreadsheetId);
        RebuildAllSheets(service, spreadsheetId, entries);
    }

    public static string GetSpreadsheetUrl()
    {
        var id = GetSpreadsheetId();
        return id != null ? $"https://docs.google.com/spreadsheets/d/{id}" : "";
    }

    private static SheetsService CreateSheetsService()
    {
        using var stream = new FileStream(CredentialPath, FileMode.Open, FileAccess.Read);
        var credential = GoogleCredential.FromStream(stream)
            .CreateScoped(SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile);

        return new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });
    }

    private static string EnsureSpreadsheetExists(SheetsService service)
    {
        var existingId = GetSpreadsheetId();
        if (existingId != null)
        {
            // Verify the spreadsheet still exists
            try
            {
                service.Spreadsheets.Get(existingId).Execute();
                return existingId;
            }
            catch
            {
                // Spreadsheet was deleted, clear stale ID and recreate
                _spreadsheetId = null;
                if (File.Exists(SpreadsheetIdFile)) File.Delete(SpreadsheetIdFile);
            }
        }

        // Create new spreadsheet with all required sheets
        var spreadsheet = new Spreadsheet
        {
            Properties = new SpreadsheetProperties { Title = SpreadsheetTitle },
            Sheets = new List<Sheet>
            {
                new Sheet { Properties = new SheetProperties { Title = "Raw Data", Hidden = true } },
                new Sheet { Properties = new SheetProperties { Title = "Deployment Hours Report", Index = 1 } },
                new Sheet { Properties = new SheetProperties { Title = "Ticket Hours Report", Index = 2 } },
                new Sheet { Properties = new SheetProperties { Title = "Contract Details", Index = 3 } },
                new Sheet { Properties = new SheetProperties { Title = "Hours Per Month", Index = 4 } }
            }
        };

        var created = service.Spreadsheets.Create(spreadsheet).Execute();
        _spreadsheetId = created.SpreadsheetId;
        File.WriteAllText(SpreadsheetIdFile, _spreadsheetId);

        // Write headers to Raw Data
        var headers = new List<object> { "Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Ticket Status", "Hours", "Date" };
        var headerUpdate = service.Spreadsheets.Values.Update(
            new ValueRange { Values = new List<IList<object>> { headers } },
            _spreadsheetId,
            "'Raw Data'!A1:H1");
        headerUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        headerUpdate.Execute();

        // Share with the user's email
        ShareSpreadsheet(_spreadsheetId);

        // Protect summary sheets
        ProtectSummarySheets(service, created);

        return _spreadsheetId;
    }

    private static void ShareSpreadsheet(string spreadsheetId)
    {
        if (!File.Exists(ShareEmailFile)) return;

        var email = File.ReadAllText(ShareEmailFile).Trim();
        if (string.IsNullOrEmpty(email)) return;

        using var stream = new FileStream(CredentialPath, FileMode.Open, FileAccess.Read);
        var credential = GoogleCredential.FromStream(stream)
            .CreateScoped(DriveService.Scope.DriveFile);

        var driveService = new DriveService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });

        var permission = new DrivePermission
        {
            Type = "user",
            Role = "writer",
            EmailAddress = email
        };

        driveService.Permissions.Create(permission, spreadsheetId).Execute();
    }

    private static void ProtectSummarySheets(SheetsService service, Spreadsheet spreadsheet)
    {
        var requests = new List<Request>();

        foreach (var sheet in spreadsheet.Sheets)
        {
            var title = sheet.Properties.Title;
            if (title == "Contract Details" || title == "Hours Per Month")
            {
                requests.Add(new Request
                {
                    AddProtectedRange = new AddProtectedRangeRequest
                    {
                        ProtectedRange = new ProtectedRange
                        {
                            Range = new GridRange { SheetId = sheet.Properties.SheetId },
                            Description = "Read-only summary sheet",
                            WarningOnly = true
                        }
                    }
                });
            }
        }

        if (requests.Count > 0)
        {
            service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests },
                spreadsheet.SpreadsheetId).Execute();
        }
    }

    private static string? GetSpreadsheetId()
    {
        if (_spreadsheetId != null) return _spreadsheetId;
        if (!File.Exists(SpreadsheetIdFile)) return null;

        var id = File.ReadAllText(SpreadsheetIdFile).Trim();
        if (string.IsNullOrEmpty(id)) return null;

        _spreadsheetId = id;
        return id;
    }

    private static List<SupportEntry> ReadAllRawData(SheetsService service, string spreadsheetId)
    {
        var entries = new List<SupportEntry>();

        var response = service.Spreadsheets.Values.Get(spreadsheetId, "'Raw Data'!A2:H").Execute();
        if (response.Values == null) return entries;

        foreach (var row in response.Values)
        {
            if (row.Count < 1 || string.IsNullOrWhiteSpace(row[0]?.ToString())) continue;
            try
            {
                var dateStr = row.Count > 7 ? row[7]?.ToString() ?? "" : "";
                DateTime date = DateTime.Today;
                if (!string.IsNullOrEmpty(dateStr))
                {
                    // Try ISO format first, then dd/MM/yyyy, then general parse
                    if (!DateTime.TryParseExact(dateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                        if (!DateTime.TryParseExact(dateStr, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                            if (!DateTime.TryParse(dateStr, out date))
                                date = DateTime.Today;
                }

                entries.Add(new SupportEntry
                {
                    ServiceType = row[0]?.ToString() ?? "",
                    TicketReference = row.Count > 1 ? row[1]?.ToString() ?? "" : "",
                    Subject = row.Count > 2 ? row[2]?.ToString() ?? "" : "",
                    HandledBy = row.Count > 3 ? row[3]?.ToString() ?? "" : "",
                    InitiatedBy = row.Count > 4 ? row[4]?.ToString() ?? "" : "",
                    TicketStatus = row.Count > 5 ? row[5]?.ToString() ?? "" : "",
                    Hours = row.Count > 6 && double.TryParse(row[6]?.ToString(), out var h) ? h : 0,
                    Date = date
                });
            }
            catch { }
        }

        return entries;
    }

    private static void RebuildAllSheets(SheetsService service, string spreadsheetId, List<SupportEntry> entries)
    {
        // Clear all derived sheets
        var clearRequest = new BatchClearValuesRequest
        {
            Ranges = new List<string>
            {
                "'Deployment Hours Report'!A1:Z",
                "'Ticket Hours Report'!A1:Z",
                "'Contract Details'!A1:Z",
                "'Hours Per Month'!A1:Z"
            }
        };
        service.Spreadsheets.Values.BatchClear(clearRequest, spreadsheetId).Execute();

        var batchData = new List<ValueRange>();

        // --- Deployment Hours Report ---
        var deployData = new List<IList<object>>
        {
            new List<object> { "Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Hours", "Date" }
        };
        foreach (var e in entries.Where(e => e.ServiceType != "Support"))
        {
            deployData.Add(new List<object>
            {
                e.ServiceType == "Engineer" ? "Deployment" : e.ServiceType,
                e.TicketReference, e.Subject, e.HandledBy, e.InitiatedBy,
                e.Hours, e.Date.ToString("yyyy-MM-dd")
            });
        }
        batchData.Add(new ValueRange { Range = "'Deployment Hours Report'!A1", Values = deployData });

        // --- Ticket Hours Report ---
        var ticketData = new List<IList<object>>
        {
            new List<object> { "Ticket", "Ticket Status", "Initiated By", "Handled By", "Date", "Hours", "Total Hours" }
        };
        var supportEntries = entries.Where(e => e.ServiceType == "Support").ToList();
        foreach (var e in supportEntries)
        {
            ticketData.Add(new List<object>
            {
                e.TicketReference, e.TicketStatus, e.InitiatedBy, e.HandledBy,
                e.Date.ToString("yyyy-MM-dd"), e.Hours
            });
        }
        batchData.Add(new ValueRange { Range = "'Ticket Hours Report'!A1", Values = ticketData });

        // Total Hours formula in G2
        int ticketLastRow = Math.Max(supportEntries.Count + 1, 2);
        batchData.Add(new ValueRange
        {
            Range = "'Ticket Hours Report'!G2",
            Values = new List<IList<object>> { new List<object> { $"=SUM(F2:F{ticketLastRow})" } }
        });

        // --- Contract Details ---
        var contractData = new List<IList<object>>
        {
            new List<object> { "Service", "Total Hours" },
            new List<object> { "Deployment", "=SUMIF('Deployment Hours Report'!A2:A1000,\"Deployment\",'Deployment Hours Report'!F2:F1000)" },
            new List<object> { "Management", "=SUMIF('Deployment Hours Report'!A2:A1000,\"Management\",'Deployment Hours Report'!F2:F1000)" },
            new List<object> { "Support", "='Ticket Hours Report'!G2" },
            new List<object> { "Total", "=SUM(B2:B4)" }
        };
        batchData.Add(new ValueRange { Range = "'Contract Details'!A1", Values = contractData });

        // --- Hours Per Month ---
        var monthData = new List<IList<object>>
        {
            new List<object> { "Month", "Management Hours", "Support Hours", "Total" }
        };

        var allMonths = entries
            .Select(e => new { e.Date.Year, e.Date.Month })
            .Distinct()
            .OrderBy(m => m.Year).ThenBy(m => m.Month)
            .ToList();

        int monthRow = 2;
        foreach (var month in allMonths)
        {
            var monthStr = $"{month.Month:D2}/{month.Year}";
            monthData.Add(new List<object>
            {
                monthStr,
                $"=SUMPRODUCT(('Deployment Hours Report'!$A$2:$A$1000=\"Management\")*(TEXT('Deployment Hours Report'!$G$2:$G$1000,\"MM/YYYY\")=A{monthRow})*'Deployment Hours Report'!$F$2:$F$1000)",
                $"=SUMPRODUCT((TEXT('Ticket Hours Report'!$E$2:$E$1000,\"MM/YYYY\")=A{monthRow})*'Ticket Hours Report'!$F$2:$F$1000)+SUMPRODUCT(('Deployment Hours Report'!$A$2:$A$1000=\"Deployment\")*(TEXT('Deployment Hours Report'!$G$2:$G$1000,\"MM/YYYY\")=A{monthRow})*'Deployment Hours Report'!$F$2:$F$1000)",
                $"=B{monthRow}+C{monthRow}"
            });
            monthRow++;
        }

        int lastDataRow = monthRow - 1;
        monthData.Add(new List<object>
        {
            "Total",
            $"=SUM(B2:B{lastDataRow})",
            $"=SUM(C2:C{lastDataRow})",
            $"=SUM(D2:D{lastDataRow})"
        });

        batchData.Add(new ValueRange { Range = "'Hours Per Month'!A1", Values = monthData });

        // Write all data in one batch
        var batchUpdate = new BatchUpdateValuesRequest
        {
            ValueInputOption = "USER_ENTERED",
            Data = batchData
        };
        service.Spreadsheets.Values.BatchUpdate(batchUpdate, spreadsheetId).Execute();

        // Apply header formatting and data validation
        ApplyFormatting(service, spreadsheetId);
    }

    private static void ApplyFormatting(SheetsService service, string spreadsheetId)
    {
        var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
        var requests = new List<Request>();

        foreach (var sheet in spreadsheet.Sheets)
        {
            var sheetId = sheet.Properties.SheetId;
            var title = sheet.Properties.Title;
            if (title == "Raw Data") continue;

            // Format header row (bold, white text, blue background)
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = sheetId, StartRowIndex = 0, EndRowIndex = 1 },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            TextFormat = new TextFormat
                            {
                                Bold = true,
                                ForegroundColor = new Google.Apis.Sheets.v4.Data.Color { Red = 1, Green = 1, Blue = 1 }
                            },
                            BackgroundColor = new Google.Apis.Sheets.v4.Data.Color { Red = 68f / 255, Green = 114f / 255, Blue = 196f / 255 },
                            HorizontalAlignment = "CENTER"
                        }
                    },
                    Fields = "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)"
                }
            });

            // Freeze header row
            requests.Add(new Request
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest
                {
                    Properties = new SheetProperties
                    {
                        SheetId = sheetId,
                        GridProperties = new GridProperties { FrozenRowCount = 1 }
                    },
                    Fields = "gridProperties.frozenRowCount"
                }
            });

            // Add data validation dropdown for Deployment Hours Report Type Service column
            if (title == "Deployment Hours Report")
            {
                requests.Add(new Request
                {
                    SetDataValidation = new SetDataValidationRequest
                    {
                        Range = new GridRange { SheetId = sheetId, StartRowIndex = 1, EndRowIndex = 1000, StartColumnIndex = 0, EndColumnIndex = 1 },
                        Rule = new DataValidationRule
                        {
                            Condition = new BooleanCondition
                            {
                                Type = "ONE_OF_LIST",
                                Values = new List<ConditionValue>
                                {
                                    new ConditionValue { UserEnteredValue = "Management" },
                                    new ConditionValue { UserEnteredValue = "Deployment" }
                                }
                            },
                            ShowCustomUi = true
                        }
                    }
                });
            }
        }

        if (requests.Count > 0)
        {
            service.Spreadsheets.BatchUpdate(
                new BatchUpdateSpreadsheetRequest { Requests = requests },
                spreadsheetId).Execute();
        }
    }
}
