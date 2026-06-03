using System;
using System.Drawing;
using System.Windows.Forms;

namespace PExL.AddIn.Templates
{
    /// <summary>
    /// Small calendar dialog for the Date Picker template. Returns the chosen
    /// date so the runner can emit a PExL date literal (#YYYY-MM-DD#).
    /// </summary>
    public sealed class DatePickerForm : Form
    {
        private readonly MonthCalendar _cal;
        public DateTime SelectedDate => _cal.SelectionStart;

        public DatePickerForm()
        {
            Text = "PExL - Pick a date";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false; MinimizeBox = false;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            _cal = new MonthCalendar { MaxSelectionCount = 1, Location = new Point(12, 12) };

            var ok = new Button { Text = "Insert", DialogResult = DialogResult.OK, Width = 90 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(0, 8, 0, 0)
            };
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);

            var layout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill
            };
            layout.Controls.Add(_cal);
            layout.Controls.Add(buttons);
            Controls.Add(layout);

            AcceptButton = ok; CancelButton = cancel;
        }
    }
}
