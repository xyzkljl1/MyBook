using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;

namespace MyBook;

public class SegmentedSelector : ListBox
{
    static SegmentedSelector()
    {
        DefaultStyleKeyProperty.OverrideMetadata(typeof(SegmentedSelector), new FrameworkPropertyMetadata(typeof(SegmentedSelector)));
    }

    protected override void OnItemsChanged(NotifyCollectionChangedEventArgs e)
    {
        base.OnItemsChanged(e);
        if (SelectedIndex < 0 && Items.Count > 0)
            SetCurrentValue(SelectedIndexProperty, 0);
    }

    protected override void OnSelectionChanged(SelectionChangedEventArgs e)
    {
        base.OnSelectionChanged(e);
        // A populated selector always retains one choice, including Ctrl-click on the selected item.
        if (SelectedIndex < 0 && Items.Count > 0)
        {
            var previousIndex = e.RemovedItems.Count > 0 ? Items.IndexOf(e.RemovedItems[0]) : 0;
            SetCurrentValue(SelectedIndexProperty, Math.Max(0, previousIndex));
        }
    }
}
