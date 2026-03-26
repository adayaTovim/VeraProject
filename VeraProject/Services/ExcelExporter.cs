using System.Diagnostics;
using ClosedXML.Excel;
using VeraProject.Models;

namespace VeraProject.Services;

public static class ExcelExporter
{
    private static readonly string FilePath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "IUCC-HOURS REPORT.xlsx");

    private const string RawDataSheet = "Raw Data";

    public static void AppendEntry(SupportEntry entry)
    {
        if (File.Exists(FilePath) && IsFileLocked(FilePath))
            ForceCloseExcel(FilePath);

        using var workbook = File.Exists(FilePath)
            ? new XLWorkbook(FilePath)
            : new XLWorkbook();

        IXLWorksheet dataSheet;

        if (workbook.Worksheets.Contains(RawDataSheet))
        {
            dataSheet = workbook.Worksheet(RawDataSheet);
        }
        else
        {
            dataSheet = workbook.Worksheets.Add(RawDataSheet);
            var headers = new[] { "Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Ticket Status", "Hours", "Date" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = dataSheet.Cell(1, i + 1);
                cell.Value = headers[i];
            }
            StyleHeader(dataSheet, 1, headers.Length);
        }

        // Append new row to raw data
        int nextRow = dataSheet.LastRowUsed()!.RowNumber() + 1;
        dataSheet.Cell(nextRow, 1).Value = entry.ServiceType;
        dataSheet.Cell(nextRow, 2).Value = entry.TicketReference;
        dataSheet.Cell(nextRow, 3).Value = entry.Subject;
        dataSheet.Cell(nextRow, 4).Value = entry.HandledBy;
        dataSheet.Cell(nextRow, 5).Value = entry.InitiatedBy;
        dataSheet.Cell(nextRow, 6).Value = entry.TicketStatus;
        dataSheet.Cell(nextRow, 7).Value = entry.Hours;
        var dateCell = dataSheet.Cell(nextRow, 8);
        dateCell.Value = entry.Date;
        dateCell.Style.DateFormat.Format = "dd/MM/yyyy";
        dataSheet.Columns().AdjustToContents();
        dataSheet.Visibility = XLWorksheetVisibility.Hidden;

        // Read all entries for summary sheets
        var entries = ReadAllEntries(dataSheet);

        // Deployment Hours Report (Management + Engineer only)
        RebuildSheet(workbook, "Deployment Hours Report", ws =>
        {
            var dHeaders = new[] { "Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Hours", "Date" };
            for (int i = 0; i < dHeaders.Length; i++)
                ws.Cell(1, i + 1).Value = dHeaders[i];
            StyleHeader(ws, 1, dHeaders.Length);

            var deployEntries = entries.Where(e => e.ServiceType != "Support").ToList();
            int row = 2;
            foreach (var e in deployEntries)
            {
                ws.Cell(row, 1).Value = e.ServiceType == "Engineer" ? "Deployment" : e.ServiceType;
                ws.Cell(row, 2).Value = e.TicketReference;
                ws.Cell(row, 3).Value = e.Subject;
                ws.Cell(row, 4).Value = e.HandledBy;
                ws.Cell(row, 5).Value = e.InitiatedBy;
                ws.Cell(row, 6).Value = e.Hours;
                ws.Cell(row, 6).Style.NumberFormat.Format = "0.00";
                var dc = ws.Cell(row, 7);
                dc.Value = e.Date;
                dc.Style.DateFormat.Format = "dd/MM/yyyy";
                row++;
            }

            StyleDataRows(ws, 2, row - 1, dHeaders.Length);
            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents();
        });

        // Contract Details
        RebuildSheet(workbook, "Contract Details", ws =>
        {
            ws.Cell(1, 1).Value = "Service";
            ws.Cell(1, 2).Value = "Total Hours";
            StyleHeader(ws, 1, 2);

            double deploymentHours = entries.Where(e => e.ServiceType == "Engineer").Sum(e => e.Hours);
            double managementHours = entries.Where(e => e.ServiceType == "Management").Sum(e => e.Hours);
            double supportHours = entries.Where(e => e.ServiceType == "Support").Sum(e => e.Hours);

            ws.Cell(2, 1).Value = "Deployment";
            ws.Cell(2, 2).Value = deploymentHours;
            ws.Cell(3, 1).Value = "Management";
            ws.Cell(3, 2).Value = managementHours;
            ws.Cell(4, 1).Value = "Support";
            ws.Cell(4, 2).Value = supportHours;

            // Total row
            ws.Cell(5, 1).Value = "Total";
            ws.Cell(5, 2).Value = deploymentHours + managementHours + supportHours;
            ws.Cell(5, 1).Style.Font.Bold = true;
            ws.Cell(5, 2).Style.Font.Bold = true;
            ws.Cell(5, 1).Style.Border.TopBorder = XLBorderStyleValues.Double;
            ws.Cell(5, 2).Style.Border.TopBorder = XLBorderStyleValues.Double;

            for (int i = 2; i <= 5; i++)
                ws.Cell(i, 2).Style.NumberFormat.Format = "0.00";

            StyleDataRows(ws, 2, 4, 2);
            ws.Columns().AdjustToContents();
        });

        // Hours Per Month (Management + Support side by side)
        RebuildSheet(workbook, "Hours Per Month", ws =>
        {
            ws.Cell(1, 1).Value = "Month";
            ws.Cell(1, 2).Value = "Management Hours";
            ws.Cell(1, 3).Value = "Support Hours";
            ws.Cell(1, 4).Value = "Total";
            StyleHeader(ws, 1, 4);

            var allMonths = entries
                .Where(e => e.ServiceType == "Management" || e.ServiceType == "Support")
                .Select(e => new { e.Date.Year, e.Date.Month })
                .Distinct()
                .OrderBy(m => m.Year).ThenBy(m => m.Month);

            int row = 2;
            double totalMgmt = 0, totalSupport = 0;
            foreach (var month in allMonths)
            {
                double mgmt = entries
                    .Where(e => e.ServiceType == "Management" && e.Date.Year == month.Year && e.Date.Month == month.Month)
                    .Sum(e => e.Hours);
                double supp = entries
                    .Where(e => e.ServiceType == "Support" && e.Date.Year == month.Year && e.Date.Month == month.Month)
                    .Sum(e => e.Hours);

                ws.Cell(row, 1).Value = $"{month.Month:D2}/{month.Year}";
                ws.Cell(row, 2).Value = mgmt;
                ws.Cell(row, 3).Value = supp;
                ws.Cell(row, 4).Value = mgmt + supp;
                totalMgmt += mgmt;
                totalSupport += supp;
                row++;
            }

            // Total row
            ws.Cell(row, 1).Value = "Total";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = totalMgmt;
            ws.Cell(row, 3).Value = totalSupport;
            ws.Cell(row, 4).Value = totalMgmt + totalSupport;
            for (int i = 1; i <= 4; i++)
            {
                ws.Cell(row, i).Style.Font.Bold = true;
                ws.Cell(row, i).Style.Border.TopBorder = XLBorderStyleValues.Double;
            }

            for (int r = 2; r <= row; r++)
                for (int c = 2; c <= 4; c++)
                    ws.Cell(r, c).Style.NumberFormat.Format = "0.00";

            StyleDataRows(ws, 2, row - 1, 4);
            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents();
        });

        // Ticket Hours Report (Support only)
        RebuildSheet(workbook, "Ticket Hours Report", ws =>
        {
            var ticketHeaders = new[] { "Ticket", "Ticket Status", "Initiated By", "Handled By", "Date", "Hours" };
            for (int i = 0; i < ticketHeaders.Length; i++)
                ws.Cell(1, i + 1).Value = ticketHeaders[i];
            StyleHeader(ws, 1, ticketHeaders.Length);

            var supportEntries = entries.Where(e => e.ServiceType == "Support").ToList();
            int row = 2;
            foreach (var e in supportEntries)
            {
                ws.Cell(row, 1).Value = e.TicketReference;
                ws.Cell(row, 2).Value = e.TicketStatus;
                ws.Cell(row, 3).Value = e.InitiatedBy;
                ws.Cell(row, 4).Value = e.HandledBy;
                var dc = ws.Cell(row, 5);
                dc.Value = e.Date;
                dc.Style.DateFormat.Format = "dd/MM/yyyy";
                ws.Cell(row, 6).Value = e.Hours;
                ws.Cell(row, 6).Style.NumberFormat.Format = "0.00";
                row++;
            }

            // Total row
            if (supportEntries.Count > 0)
            {
                ws.Cell(row, 5).Value = "Total";
                ws.Cell(row, 5).Style.Font.Bold = true;
                ws.Cell(row, 6).Value = supportEntries.Sum(e => e.Hours);
                ws.Cell(row, 6).Style.Font.Bold = true;
                ws.Cell(row, 6).Style.NumberFormat.Format = "0.00";
                ws.Cell(row, 5).Style.Border.TopBorder = XLBorderStyleValues.Double;
                ws.Cell(row, 6).Style.Border.TopBorder = XLBorderStyleValues.Double;
            }

            StyleDataRows(ws, 2, row - 1, ticketHeaders.Length);
            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents();
        });

        workbook.SaveAs(FilePath);
    }

    private static List<SupportEntry> ReadAllEntries(IXLWorksheet dataSheet)
    {
        var entries = new List<SupportEntry>();
        int lastRow = dataSheet.LastRowUsed()?.RowNumber() ?? 1;

        for (int row = 2; row <= lastRow; row++)
        {
            entries.Add(new SupportEntry
            {
                ServiceType = dataSheet.Cell(row, 1).GetString(),
                TicketReference = dataSheet.Cell(row, 2).GetString(),
                Subject = dataSheet.Cell(row, 3).GetString(),
                HandledBy = dataSheet.Cell(row, 4).GetString(),
                InitiatedBy = dataSheet.Cell(row, 5).GetString(),
                TicketStatus = dataSheet.Cell(row, 6).GetString(),
                Hours = dataSheet.Cell(row, 7).GetDouble(),
                Date = dataSheet.Cell(row, 8).GetDateTime()
            });
        }

        return entries;
    }

    private static void RebuildSheet(IXLWorkbook workbook, string sheetName, Action<IXLWorksheet> build)
    {
        if (workbook.Worksheets.Contains(sheetName))
            workbook.Worksheets.Delete(sheetName);

        var ws = workbook.Worksheets.Add(sheetName);
        build(ws);
    }

    private static void StyleHeader(IXLWorksheet ws, int row, int colCount)
    {
        for (int i = 1; i <= colCount; i++)
        {
            var cell = ws.Cell(row, i);
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(68, 114, 196);
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
    }

    private static void StyleDataRows(IXLWorksheet ws, int startRow, int endRow, int colCount)
    {
        for (int row = startRow; row <= endRow; row++)
        {
            for (int col = 1; col <= colCount; col++)
            {
                var cell = ws.Cell(row, col);
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(200, 200, 200);
            }

            // Alternating row colors
            if (row % 2 == 0)
            {
                var range = ws.Range(row, 1, row, colCount);
                range.Style.Fill.BackgroundColor = XLColor.FromArgb(234, 240, 248);
            }
        }
    }

    private static bool IsFileLocked(string path)
    {
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
    }

    private static void ForceCloseExcel(string path)
    {
        string fileName = Path.GetFileName(path);
        var excelProcesses = Process.GetProcessesByName("EXCEL");

        foreach (var process in excelProcesses)
        {
            try
            {
                process.Kill();
                process.WaitForExit(3000);
            }
            catch { }
        }

        // Wait for file to be released
        for (int i = 0; i < 10; i++)
        {
            if (!IsFileLocked(path)) return;
            Thread.Sleep(300);
        }

        throw new IOException("Could not release the Excel file. Please close it manually.");
    }
}
