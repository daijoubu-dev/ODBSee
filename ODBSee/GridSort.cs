using System.Globalization;

namespace ODBSee
{
    /// <summary>
    /// Type-aware sorting for the result cache. Values are stored as strings (cheap to fetch and display),
    /// so for each sort we look at the column once: if every non-empty value is a number it sorts numerically,
    /// otherwise if every non-empty value is a date/time it sorts chronologically, otherwise as text.
    /// Empty values always go last. The sort is stable, so sorting by B then A keeps B's order within A.
    /// </summary>
    public static class GridSort
    {
        public enum Kind { Number, Date, Text }

        private const NumberStyles NumStyles = NumberStyles.Float | NumberStyles.AllowThousands;
        private const DateTimeStyles DateStyles = DateTimeStyles.AllowWhiteSpaces;

        public static List<string[]> Sort(List<string[]> data, int col, bool ascending) => Sort(data, col, ascending, out _);

        public static List<string[]> Sort(List<string[]> data, int col, bool ascending, out Kind kind)
        {
            int n = data.Count;

            // Split out empties (always sorted to the bottom, original order kept).
            var nonEmpty = new List<int>(n);
            var empty = new List<int>();
            for (int i = 0; i < n; i++)
            {
                var row = data[i];
                if (col < row.Length && !string.IsNullOrWhiteSpace(row[col])) nonEmpty.Add(i);
                else empty.Add(i);
            }

            int[] order;
            var keys = new double[nonEmpty.Count];

            if (TryKeys(data, col, nonEmpty, keys, TryNumber))
            {
                kind = Kind.Number;
                order = SortByKeys(nonEmpty, keys, ascending);
            }
            else if (TryKeys(data, col, nonEmpty, keys, TryDate))
            {
                kind = Kind.Date;
                order = SortByKeys(nonEmpty, keys, ascending);
            }
            else
            {
                kind = Kind.Text;
                order = nonEmpty.ToArray();
                int sign = ascending ? 1 : -1;
                Array.Sort(order, (a, b) =>
                {
                    int c = StringComparer.OrdinalIgnoreCase.Compare(data[a][col], data[b][col]);
                    if (c == 0) c = string.CompareOrdinal(data[a][col], data[b][col]);
                    return c != 0 ? c * sign : a.CompareTo(b); // index tiebreak keeps it stable
                });
            }

            var result = new List<string[]>(n);
            foreach (var i in order) result.Add(data[i]);
            foreach (var i in empty) result.Add(data[i]);
            return result;
        }

        private delegate bool KeyParser(string s, out double key);

        private static bool TryKeys(List<string[]> data, int col, List<int> rows, double[] keys, KeyParser parse)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                if (!parse(data[rows[i]][col], out keys[i])) return false;
            }
            return true;
        }

        private static bool TryNumber(string s, out double key)
        {
            return double.TryParse(s, NumStyles, CultureInfo.CurrentCulture, out key)
                || double.TryParse(s, NumStyles, CultureInfo.InvariantCulture, out key);
        }

        private static bool TryDate(string s, out double key)
        {
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateStyles, out var d) ||
                DateTime.TryParse(s, CultureInfo.InvariantCulture, DateStyles, out d))
            {
                key = d.Ticks;
                return true;
            }
            if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateStyles, out var dto))
            {
                key = dto.UtcTicks;
                return true;
            }
            if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var ts))
            {
                key = ts.Ticks;
                return true;
            }
            key = 0;
            return false;
        }

        private static int[] SortByKeys(List<int> rows, double[] keys, bool ascending)
        {
            var pos = new int[rows.Count];
            for (int i = 0; i < pos.Length; i++) pos[i] = i;
            int sign = ascending ? 1 : -1;
            Array.Sort(pos, (a, b) =>
            {
                int c = keys[a].CompareTo(keys[b]);
                return c != 0 ? c * sign : a.CompareTo(b);
            });

            var order = new int[pos.Length];
            for (int i = 0; i < pos.Length; i++) order[i] = rows[pos[i]];
            return order;
        }
    }
}
