using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MyBook;

public sealed class AssetHistoryChart : Grid
{
    public static readonly DependencyProperty PointsProperty = DependencyProperty.Register(
        nameof(Points), typeof(IReadOnlyList<AssetHistoryPoint>), typeof(AssetHistoryChart),
        new PropertyMetadata(null, (sender, _) => ((AssetHistoryChart)sender).Draw()));

    public IReadOnlyList<AssetHistoryPoint>? Points
    {
        get => (IReadOnlyList<AssetHistoryPoint>?)GetValue(PointsProperty);
        set => SetValue(PointsProperty, value);
    }

    readonly Canvas canvas = new() { Background = Brushes.Transparent, ClipToBounds = true };
    readonly TextBox selected = Label("");
    readonly Line hoverLine = new()
    {
        Stroke = Brushes.SlateGray, StrokeThickness = 1, StrokeDashArray = new() { 4, 3 },
        IsHitTestVisible = false, Visibility = Visibility.Collapsed
    };
    const double Left = 95, Right = 12, Top = 12, Bottom = 28;

    public AssetHistoryChart()
    {
        RowDefinitions.Add(new() { Height = GridLength.Auto });
        RowDefinitions.Add(new());
        selected.Margin = new(Left, 0, 0, 4);
        selected.FontWeight = FontWeights.SemiBold;
        Children.Add(selected);
        SetRow(canvas, 1);
        Children.Add(canvas);
        canvas.SizeChanged += (_, _) => Draw();
        canvas.MouseMove += SelectNearest;
        canvas.MouseLeave += (_, _) => hoverLine.Visibility = Visibility.Collapsed;
    }

    static TextBox Label(string text) => new()
    {
        Text = text, IsReadOnly = true, IsReadOnlyCaretVisible = true,
        BorderThickness = new(0), Padding = new(0), Background = Brushes.Transparent,
        Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)), FontSize = 12
    };

    void SelectNearest(object sender, MouseEventArgs args)
        => UpdateHover(args.GetPosition(canvas));

    void UpdateHover(Point position)
    {
        hoverLine.Visibility = Visibility.Collapsed;
        if (Points is not { Count: > 0 } points || hoverLine.Parent != canvas
            || canvas.ActualWidth <= Left + Right || canvas.ActualHeight <= Top + Bottom
            || position.X < Left || position.X > canvas.ActualWidth - Right
            || position.Y < Top || position.Y > canvas.ActualHeight - Bottom) return;
        var width = canvas.ActualWidth - Left - Right;
        var fraction = (position.X - Left) / width;
        var date = points[0].Date.AddDays((points[^1].Date - points[0].Date).TotalDays * fraction);
        var nearest = points.MinBy(point => Math.Abs((point.Date - date).TotalDays))!;
        Show(nearest);
        hoverLine.X1 = hoverLine.X2 = Left + (nearest.Date - points[0].Date).TotalDays
            / Math.Max(1, (points[^1].Date - points[0].Date).TotalDays) * width;
        hoverLine.Y1 = Top;
        hoverLine.Y2 = canvas.ActualHeight - Bottom;
        hoverLine.Visibility = Visibility.Visible;
    }

    void Show(AssetHistoryPoint point) => selected.Text = point.TotalRmb is decimal amount
        ? $"{point.Date:yyyy-MM-dd}   ¥{amount:N2}" : $"{point.Date:yyyy-MM-dd}   无数据";

    void Draw()
    {
        canvas.Children.Clear();
        hoverLine.Visibility = Visibility.Collapsed;
        var available = Points?.Where(point => point.TotalRmb.HasValue).ToList() ?? [];
        if (available.Count == 0) { selected.Text = "暂无资金历史"; return; }
        Show(Points![^1]);
        var width = canvas.ActualWidth - Left - Right;
        var height = canvas.ActualHeight - Top - Bottom;
        if (width <= 0 || height <= 0) return;
        var min = available.Min(point => point.TotalRmb!.Value);
        var max = available.Max(point => point.TotalRmb!.Value);
        var padding = Math.Max((max - min) * 0.08m, Math.Max(Math.Abs(max) * 0.001m, 1m));
        min -= padding;
        max += padding;
        var days = Math.Max(1, (Points![^1].Date - Points[0].Date).TotalDays);
        double X(AssetHistoryPoint point) => Left + (point.Date - Points[0].Date).TotalDays / days * width;
        double Y(decimal value) => Top + (double)((max - value) / (max - min)) * height;
        for (var i = 0; i <= 3; i++)
        {
            var value = min + (max - min) * i / 3;
            var y = Y(value);
            canvas.Children.Add(new Line { X1 = Left, X2 = Left + width, Y1 = y, Y2 = y,
                Stroke = new SolidColorBrush(Color.FromRgb(226, 232, 240)), StrokeThickness = 1 });
            AddLabel($"¥{value:N0}", 0, y - 8, Left - 10, TextAlignment.Right);
        }
        for (var i = 0; i <= 2; i++)
        {
            var date = Points[0].Date.AddDays((Points[^1].Date - Points[0].Date).TotalDays * i / 2);
            AddLabel(date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                Left + width * i / 2 - (i == 2 ? 82 : i == 1 ? 41 : 0), Top + height + 8, 82);
        }
        Polyline? line = null;
        foreach (var point in Points)
        {
            if (point.TotalRmb is not decimal value) { line = null; continue; }
            if (line is null)
            {
                line = new Polyline { Stroke = Brushes.Teal, StrokeThickness = 2, IsHitTestVisible = false };
                canvas.Children.Add(line);
            }
            line.Points.Add(new(X(point), Y(value)));
        }
        canvas.Children.Add(hoverLine);
        if (canvas.IsMouseOver) UpdateHover(Mouse.GetPosition(canvas));
    }

    void AddLabel(string text, double x, double y, double width, TextAlignment alignment = TextAlignment.Left)
    {
        var label = Label(text);
        label.Width = width;
        label.TextAlignment = alignment;
        Canvas.SetLeft(label, x);
        Canvas.SetTop(label, y);
        canvas.Children.Add(label);
    }
}
