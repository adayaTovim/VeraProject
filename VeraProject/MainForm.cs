using System.Diagnostics;
using VeraProject.Models;
using VeraProject.Services;

namespace VeraProject;

public class MainForm : Form
{
    private readonly ComboBox _cmbServiceType;
    private readonly TextBox _txtTicketRef;
    private readonly TextBox _txtSubject;
    private readonly TextBox _txtHandledBy;
    private readonly TextBox _txtInitiatedBy;
    private readonly NumericUpDown _numHours;
    private readonly DateTimePicker _dtpDate;

    public MainForm()
    {
        Text = "Deployment Hours Support";
        AutoScaleMode = AutoScaleMode.Font;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 6,
            Padding = new Padding(10),
            AutoSize = true
        };

        // Column widths: Label, Control, Label, Control
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));

        for (int i = 0; i < 6; i++)
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Row 0: Type Service | Ticket Reference
        layout.Controls.Add(CreateLabel("Type Service:"), 0, 0);
        _cmbServiceType = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbServiceType.Items.AddRange(new object[] { "Management", "Engineer", "Support" });
        _cmbServiceType.SelectedIndex = 0;
        _cmbServiceType.SelectedIndexChanged += (s, e) =>
        {
            string selected = _cmbServiceType.SelectedItem?.ToString() ?? "";
            _txtInitiatedBy.Enabled = selected == "Management" || selected == "Support";
            if (!_txtInitiatedBy.Enabled) _txtInitiatedBy.Clear();
        };
        layout.Controls.Add(_cmbServiceType, 1, 0);

        layout.Controls.Add(CreateLabel("Ticket Reference:"), 2, 0);
        _txtTicketRef = new TextBox { Dock = DockStyle.Fill };
        layout.Controls.Add(_txtTicketRef, 3, 0);

        // Row 1: Subject | Handled By
        layout.Controls.Add(CreateLabel("Subject:"), 0, 1);
        _txtSubject = new TextBox { Dock = DockStyle.Fill };
        layout.Controls.Add(_txtSubject, 1, 1);

        layout.Controls.Add(CreateLabel("Handled By:"), 2, 1);
        _txtHandledBy = new TextBox { Dock = DockStyle.Fill };
        layout.Controls.Add(_txtHandledBy, 3, 1);

        // Row 2: Initiated By | Hours
        layout.Controls.Add(CreateLabel("Initiated By:"), 0, 2);
        _txtInitiatedBy = new TextBox { Dock = DockStyle.Fill };
        layout.Controls.Add(_txtInitiatedBy, 1, 2);

        layout.Controls.Add(CreateLabel("Hours:"), 2, 2);
        _numHours = new NumericUpDown
        {
            Dock = DockStyle.Fill,
            Minimum = 0.25m,
            Maximum = 24,
            DecimalPlaces = 2,
            Increment = 0.25m,
            Value = 1
        };
        layout.Controls.Add(_numHours, 3, 2);

        // Row 3: Date
        layout.Controls.Add(CreateLabel("Date:"), 0, 3);
        _dtpDate = new DateTimePicker
        {
            Dock = DockStyle.Fill,
            Format = DateTimePickerFormat.Short
        };
        layout.Controls.Add(_dtpDate, 1, 3);

        // Row 4: Button
        var btnAdd = new Button
        {
            Text = "Add Entry",
            Width = 120,
            Height = 35,
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(0, 10, 0, 0)
        };
        btnAdd.Click += BtnAdd_Click;
        layout.Controls.Add(btnAdd, 0, 4);
        layout.SetColumnSpan(btnAdd, 2);

        Controls.Add(layout);

        // Set form size after layout is added
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
    }

    private Label CreateLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            TextAlign = ContentAlignment.MiddleRight,
            Anchor = AnchorStyles.Right,
            Margin = new Padding(3, 8, 3, 3)
        };
    }

    private void BtnAdd_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_txtSubject.Text))
        {
            MessageBox.Show("Please enter a subject.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_txtHandledBy.Text))
        {
            MessageBox.Show("Please enter handled by.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var entry = new SupportEntry
        {
            ServiceType = _cmbServiceType.SelectedItem?.ToString() ?? "Management",
            TicketReference = _txtTicketRef.Text.Trim(),
            Subject = _txtSubject.Text.Trim(),
            HandledBy = _txtHandledBy.Text.Trim(),
            InitiatedBy = _txtInitiatedBy.Text.Trim(),
            Hours = (double)_numHours.Value,
            Date = _dtpDate.Value.Date
        };

        try
        {
            // Save to local Excel
            ExcelExporter.AppendEntry(entry);

            // Save to Google Sheets and open in browser
            if (GoogleSheetsExporter.IsConfigured)
            {
                try
                {
                    GoogleSheetsExporter.AppendEntry(entry);
                    var url = GoogleSheetsExporter.GetSpreadsheetUrl();
                    if (!string.IsNullOrEmpty(url))
                        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception gsEx)
                {
                    MessageBox.Show($"Saved locally, but Google Sheets sync failed: {gsEx.Message}",
                        "Google Sheets Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            MessageBox.Show("Entry added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            _cmbServiceType.SelectedIndex = 0;
            _txtTicketRef.Clear();
            _txtSubject.Clear();
            _txtHandledBy.Clear();
            _txtInitiatedBy.Clear();
            _numHours.Value = 1;
            _dtpDate.Value = DateTime.Today;
            _txtSubject.Focus();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
