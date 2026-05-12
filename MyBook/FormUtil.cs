using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace MyBook
{
    public static class FormUtil
    {
        public class FormTable
        {
            public string Title { get; set; } = "";
            public List<string> Headers { get; set; } = new();
            public List<List<string>> Rows { get; set; } = new();
        }

        public static List<FormTable> ReadFromHTML(string html)
        {
            var forms = new List<FormTable>();
            if (String.IsNullOrWhiteSpace(html))
                return forms;

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var titleNodes = doc.DocumentNode.SelectNodes("//td[contains(concat(' ', normalize-space(@class), ' '), ' title3 ')]");
            if (titleNodes is null)
                return forms;

            foreach (var titleNode in titleNodes)
            {
                var title = CleanText(titleNode.InnerText);
                if (String.IsNullOrEmpty(title))
                    continue;

                var titleTable = titleNode.Ancestors("table").FirstOrDefault();
                var dataTable = titleTable?.SelectSingleNode("following::table[1]")
                    ?? titleNode.SelectSingleNode("following::table[1]");
                if (dataTable is null)
                    continue;

                var rows = ReadTableRows(dataTable);
                if (rows.Count == 0)
                    continue;

                forms.Add(new FormTable
                {
                    Title = title,
                    Headers = rows[0],
                    Rows = rows.Skip(1).ToList()
                });
            }

            return forms;
        }

        private static List<List<string>> ReadTableRows(HtmlNode table)
        {
            var rows = new List<List<string>>();
            var rowSpans = new Dictionary<int, RowSpanCell>();
            var maxColumns = 0;

            foreach (var tr in table.SelectNodes("./tr|./thead/tr|./tbody/tr|./tfoot/tr") ?? Enumerable.Empty<HtmlNode>())
            {
                var row = new List<string>();
                var col = 0;
                var cells = tr.ChildNodes
                    .Where(node => node.Name.Equals("td", StringComparison.OrdinalIgnoreCase)
                        || node.Name.Equals("th", StringComparison.OrdinalIgnoreCase));

                foreach (var cell in cells)
                {
                    while (rowSpans.ContainsKey(col))
                    {
                        AddRowSpanValue(row, rowSpans, col);
                        col++;
                    }

                    var colspan = ReadSpan(cell, "colspan");
                    var rowspan = ReadSpan(cell, "rowspan");
                    var text = CleanText(cell.InnerText);

                    for (var i = 0; i < colspan; i++)
                    {
                        var cellText = i == 0 ? text : "";
                        row.Add(cellText);
                        if (rowspan > 1)
                            rowSpans[col] = new RowSpanCell(cellText, rowspan - 1);
                        col++;
                    }
                }

                while (rowSpans.ContainsKey(col))
                {
                    AddRowSpanValue(row, rowSpans, col);
                    col++;
                }

                maxColumns = Math.Max(maxColumns, row.Count);
                rows.Add(row);
            }

            foreach (var row in rows)
            {
                while (row.Count < maxColumns)
                    row.Add("");
            }

            return rows;
        }

        private static void AddRowSpanValue(List<string> row, Dictionary<int, RowSpanCell> rowSpans, int col)
        {
            var rowSpan = rowSpans[col];
            row.Add(rowSpan.Text);
            rowSpan.RemainingRows--;
            if (rowSpan.RemainingRows <= 0)
                rowSpans.Remove(col);
        }

        private static int ReadSpan(HtmlNode cell, string attributeName)
        {
            var value = cell.GetAttributeValue(attributeName, "1");
            return Int32.TryParse(value, out var span) && span > 0 ? span : 1;
        }

        private static string CleanText(string text)
        {
            var decoded = (HtmlEntity.DeEntitize(text) ?? "").Replace('\u00a0', ' ');
            return Regex.Replace(decoded, @"\s+", " ").Trim();
        }

        private class RowSpanCell
        {
            public RowSpanCell(string text, int remainingRows)
            {
                Text = text;
                RemainingRows = remainingRows;
            }

            public string Text { get; }
            public int RemainingRows { get; set; }
        }
    }
}
