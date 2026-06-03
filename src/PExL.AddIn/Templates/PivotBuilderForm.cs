using System.Drawing;
using System.Windows.Forms;

namespace PExL.AddIn.Templates
{
    /// <summary>
    /// A lightweight pivot wizard: pick the group-by range, the value range, an
    /// aggregation and where to drop the result. Produces teach-as-you-go PExL
    /// (UNIQUE of the groups + a conditional aggregate filled beside them).
    /// </summary>
    public sealed class PivotBuilderForm : Form
    {
        private readonly TextBox _group = new() { Text = "A2:A1000" };
        private readonly TextBox _value = new() { Text = "C2:C1000" };
        private readonly ComboBox _agg = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _outGroups = new() { Text = "E2" };
        private readonly TextBox _outValues = new() { Text = "F2:F1000" };

        public PivotBuilderForm()
        {
            Text = "PExL - Pivot Generator";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false; MinimizeBox = false;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(14);

            _agg.Items.AddRange(new object[] { "sum", "avg", "count", "min", "max" });
            _agg.SelectedIndex = 0;

            var grid = new TableLayoutPanel
            {
                ColumnCount = 2, AutoSize = true, Dock = DockStyle.Top, ColumnStyles =
                {
                    new ColumnStyle(SizeType.Absolute, 130),
                    new ColumnStyle(SizeType.Absolute, 200)
                }
            };
            void Row(string label, Control c) { grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 0, 0) }); c.Width = 190; grid.Controls.Add(c); }

            Row("Group by (range):", _group);
            Row("Values (range):", _value);
            Row("Aggregation:", _agg);
            Row("Groups output:", _outGroups);
            Row("Values output:", _outValues);

            var ok = new Button { Text = "Build", DialogResult = DialogResult.OK, Width = 90 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom,
                AutoSize = true, Padding = new Padding(0, 10, 0, 0)
            };
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);

            var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill };
            layout.Controls.Add(grid);
            layout.Controls.Add(buttons);
            Controls.Add(layout);

            AcceptButton = ok; CancelButton = cancel;
        }

        /// <summary>The PExL the wizard generates from the current inputs.</summary>
        public string BuildPExL()
        {
            string agg = _agg.SelectedItem?.ToString() ?? "sum";
            string groups = _group.Text.Trim();
            string values = _value.Text.Trim();
            string outG = _outGroups.Text.Trim();
            string outV = _outValues.Text.Trim();
            string anchor = FirstCell(outG);

            string valueLine = agg switch
            {
                "count" => $"countWhere({groups} = {anchor}) -> {outV}",
                "avg"   => $"avgWhere({values}, {groups} = {anchor}) -> {outV}",
                "min"   => $"// min-by-group needs MINIFS:\nraw(\"MINIFS\", {values}, {groups}, {anchor}) -> {outV}",
                "max"   => $"// max-by-group needs MAXIFS:\nraw(\"MAXIFS\", {values}, {groups}, {anchor}) -> {outV}",
                _       => $"sumWhere({values}, {groups} = {anchor}) -> {outV}",
            };

            return
                "// Pivot built by the PExL wizard\n" +
                $"// Distinct {agg} of {values} grouped by {groups}\n" +
                $"unique({groups}) -> {outG}\n" +
                valueLine;
        }

        private static string FirstCell(string addr)
        {
            int colon = addr.IndexOf(':');
            return colon >= 0 ? addr.Substring(0, colon) : addr;
        }
    }
}
