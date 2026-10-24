using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.AddIn;

/// <summary>
/// The task pane: the user's choices, plus a table of the numbers behind the chart.
/// Every change recalculates straight away; there is no "Apply" button.
/// The layout is built in code, so the class opts out of the Visual Studio designer.
/// </summary>
[ComVisible(true)]
[System.ComponentModel.DesignerCategory("")]
public class DashboardPane : UserControl
{
    private readonly ComboBox _metric = NewCombo();
    private readonly ComboBox _segment = NewCombo();
    private readonly ComboBox _country = NewCombo();
    private readonly RadioButton _monthly = new() { Text = "Monthly", AutoSize = true, Checked = true };
    private readonly RadioButton _quarterly = new() { Text = "Quarterly", AutoSize = true };
    private readonly NumericUpDown _periods = new() { Minimum = 1, Maximum = 30, Value = 12, Width = 60 };
    private readonly Button _reload = new() { Text = "Reload data", AutoSize = true };
    private readonly Label _status = new() { AutoSize = true, MaximumSize = new Size(320, 0), Padding = new Padding(0, 6, 0, 6) };
    private readonly ListView _preview = new()
    {
        View = View.Details,
        FullRowSelect = true,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
        ShowItemToolTips = true,
        Dock = DockStyle.Fill
    };

    private FinancialData? _data;
    private bool _populating;

    public DashboardPane()
    {
        BuildLayout();

        _metric.SelectedIndexChanged += (_, _) => Recalculate();
        _segment.SelectedIndexChanged += (_, _) => Recalculate();
        _country.SelectedIndexChanged += (_, _) => Recalculate();
        _monthly.CheckedChanged += (_, _) => OnFrequencyChanged();
        _periods.ValueChanged += (_, _) => Recalculate();
        _reload.Click += (_, _) => Reload();
    }

    /// <summary>Where the data comes from; set by <see cref="TaskPaneManager"/>.</summary>
    public Func<FinancialData>? DataSource { get; set; }

    /// <summary>Raised after every successful calculation, with the series to chart.</summary>
    public event EventHandler<SeriesResult>? SeriesCalculated;

    /// <summary>Reads the workbook again and refreshes the choices, keeping the current ones where possible.</summary>
    public void Reload()
    {
        if (DataSource is null) return;

        try
        {
            _data = DataSource();
        }
        catch (Exception ex) when (ex is InvalidOperationException or FormatException)
        {
            _data = null;
            _preview.Items.Clear();
            ShowStatus(ex.Message, isError: true);
            return;
        }

        _populating = true;
        try
        {
            var metric = (_metric.SelectedItem as Choice)?.Value ?? "Revenue";
            Fill(_metric, _data.Metrics.Select(m => new Choice(m.Code, m.DisplayName)), metric);
            Fill(_segment, WithAll(_data.Segments), (_segment.SelectedItem as Choice)?.Value ?? Selection.All);
            Fill(_country, WithAll(_data.Countries), (_country.SelectedItem as Choice)?.Value ?? Selection.All);
            UpdatePeriodLimit();
        }
        finally
        {
            _populating = false;
        }

        Recalculate();
    }

    private void OnFrequencyChanged()
    {
        UpdatePeriodLimit();
        Recalculate();
    }

    /// <summary>The most periods that exist: every month, or every quarter the months fall in.</summary>
    private void UpdatePeriodLimit()
    {
        if (_data is null) return;

        var max = _monthly.Checked
            ? _data.Periods.Count
            : _data.Periods.Select(p => (p.Year, p.Quarter)).Distinct().Count();
        _periods.Maximum = max;
    }

    private Selection? CurrentSelection()
    {
        if (_metric.SelectedItem is not Choice metric ||
            _segment.SelectedItem is not Choice segment ||
            _country.SelectedItem is not Choice country)
            return null;

        return new Selection(
            metric.Value,
            segment.Value,
            country.Value,
            _monthly.Checked ? Frequency.Monthly : Frequency.Quarterly,
            (int)_periods.Value);
    }

    private void Recalculate()
    {
        if (_populating || _data is null || CurrentSelection() is not Selection selection) return;

        SeriesResult result;
        try
        {
            result = SeriesCalculator.Calculate(_data, selection);
        }
        catch (ArgumentException ex)
        {
            _preview.Items.Clear();
            ShowStatus(ex.Message, isError: true);
            return;
        }

        ShowPreview(result);
        SeriesCalculated?.Invoke(this, result);
    }

    private void ShowPreview(SeriesResult result)
    {
        _preview.BeginUpdate();
        _preview.Items.Clear();
        foreach (var point in result.Points)
        {
            var item = new ListViewItem([point.Label, Format(point.Value, result.Metric.IsPercent), point.IsComplete ? "" : "⚠"]);
            if (!point.IsComplete)
            {
                item.ToolTipText = string.Join(Environment.NewLine, point.Notes);
                item.ForeColor = Color.DarkOrange;
            }
            _preview.Items.Add(item);
        }
        _preview.EndUpdate();

        var warnings = result.Points.Count(p => !p.IsComplete);
        var unit = result.Metric.IsPercent ? "Percent" : "USD thousands";
        ShowStatus(warnings == 0
            ? $"{result.Metric.DisplayName} · {unit}"
            : $"{result.Metric.DisplayName} · {unit}\n⚠ {warnings} period(s) have missing data. Hover a row for details.",
            isError: false);
    }

    internal static string Format(decimal? value, bool isPercent) => value switch
    {
        null => "n/a",
        decimal v when isPercent => v.ToString("0.0", CultureInfo.InvariantCulture) + "%",
        decimal v => v.ToString("#,##0", CultureInfo.InvariantCulture)
    };

    internal void ShowError(string message) => ShowStatus(message, isError: true);

    private void ShowStatus(string message, bool isError)
    {
        _status.Text = message;
        _status.ForeColor = isError ? Color.Firebrick : SystemColors.ControlText;
    }

    private void BuildLayout()
    {
        var frequency = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        frequency.Controls.AddRange([_monthly, _quarterly]);

        var periods = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        periods.Controls.AddRange([_periods, new Label { Text = "periods, ending at the latest month", AutoSize = true, Padding = new Padding(0, 6, 0, 0) }]);

        _preview.Columns.Add("Period", 70);
        _preview.Columns.Add("Value", 100, HorizontalAlignment.Right);
        _preview.Columns.Add("", 30);

        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(8) };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddRow(grid, "Metric", _metric);
        AddRow(grid, "Segment", _segment);
        AddRow(grid, "Country", _country);
        AddRow(grid, "Periods", frequency);
        AddRow(grid, "Show", periods);
        AddRow(grid, "", _reload);
        AddFullRow(grid, _status, SizeType.AutoSize);
        AddFullRow(grid, _preview, SizeType.Percent);

        Controls.Add(grid);
    }

    private static void AddRow(TableLayoutPanel grid, string label, Control control)
    {
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.Controls.Add(new Label { Text = label, AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, 0, grid.RowCount);
        grid.Controls.Add(control, 1, grid.RowCount);
        grid.RowCount++;
    }

    private static void AddFullRow(TableLayoutPanel grid, Control control, SizeType sizeType)
    {
        grid.RowStyles.Add(sizeType == SizeType.Percent ? new RowStyle(SizeType.Percent, 100) : new RowStyle(sizeType));
        grid.Controls.Add(control, 0, grid.RowCount);
        grid.SetColumnSpan(control, 2);
        grid.RowCount++;
    }

    private static ComboBox NewCombo() => new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };

    private static IEnumerable<Choice> WithAll(IEnumerable<string> values) =>
        new[] { new Choice(Selection.All, Selection.All) }.Concat(values.Select(v => new Choice(v, v)));

    private static void Fill(ComboBox combo, IEnumerable<Choice> choices, string selected)
    {
        combo.Items.Clear();
        foreach (var choice in choices)
            combo.Items.Add(choice);

        combo.SelectedItem = combo.Items.Cast<Choice>().FirstOrDefault(c => c.Value == selected) ?? combo.Items[0];
    }

    /// <summary>A dropdown entry: the value the calculator uses and the text the user sees.</summary>
    private sealed class Choice(string value, string text)
    {
        public string Value { get; } = value;

        public override string ToString() => text;
    }
}
