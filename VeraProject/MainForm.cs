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
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Size = new Size(700, 300);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimumSize = new Size(700, 300);

        int labelWidth = 140;
        int controlWidth = 170;
        int rowHeight = 30;
        int col1X = 10, col2X = 340;

        // Row 1
        Controls.Add(CreateLabel("Type Service:", col1X, 10, labelWidth));
        _cmbServiceType = new ComboBox
        {
            Location = new Point(col1X + labelWidth, 10),
            Width = controlWidth,
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
        Controls.Add(_cmbServiceType);

        Controls.Add(CreateLabel("Ticket Reference:", col2X, 10, labelWidth));
        _txtTicketRef = new TextBox { Location = new Point(col2X + labelWidth, 10), Width = controlWidth };
        Controls.Add(_txtTicketRef);

        // Row 2
        int row2Y = 10 + rowHeight + 10;
        Controls.Add(CreateLabel("Subject:", col1X, row2Y, labelWidth));
        _txtSubject = new TextBox { Location = new Point(col1X + labelWidth, row2Y), Width = controlWidth };
        Controls.Add(_txtSubject);

        Controls.Add(CreateLabel("Handled By:", col2X, row2Y, labelWidth));
        _txtHandledBy = new TextBox { Location = new Point(col2X + labelWidth, row2Y), Width = controlWidth };
        Controls.Add(_txtHandledBy);

        // Row 3
        int row3Y = row2Y + rowHeight + 10;
        Controls.Add(CreateLabel("Initiated By:", col1X, row3Y, labelWidth));
        _txtInitiatedBy = new TextBox { Location = new Point(col1X + labelWidth, row3Y), Width = controlWidth };
        Controls.Add(_txtInitiatedBy);

        Controls.Add(CreateLabel("Hours:", col2X, row3Y, labelWidth));
        _numHours = new NumericUpDown
        {
            Location = new Point(col2X + labelWidth, row3Y),
            Width = 80,
            Minimum = 0.25m,
            Maximum = 24,
            DecimalPlaces = 2,
            Increment = 0.25m,
            Value = 1
        };
        Controls.Add(_numHours);

        // Row 4
        int row4Y = row3Y + rowHeight + 10;
        Controls.Add(CreateLabel("Date:", col1X, row4Y, labelWidth));
        _dtpDate = new DateTimePicker
        {
            Location = new Point(col1X + labelWidth, row4Y),
            Width = controlWidth,
            Format = DateTimePickerFormat.Short
        };
        Controls.Add(_dtpDate);

        // Row 5 - Button
        int row5Y = row4Y + rowHeight + 20;

        var btnAdd = new Button
        {
            Text = "Add Entry",
            Location = new Point(col1X, row5Y),
            Width = 120,
            Height = 32,
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnAdd.Click += BtnAdd_Click;
        Controls.Add(btnAdd);
    }

    private Label CreateLabel(string text, int x, int y, int width)
    {
        return new Label
        {
            Text = text,
            Location = new Point(x, y + 3),
            Width = width,
            TextAlign = ContentAlignment.MiddleRight
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
            ExcelExporter.AppendEntry(entry);
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
