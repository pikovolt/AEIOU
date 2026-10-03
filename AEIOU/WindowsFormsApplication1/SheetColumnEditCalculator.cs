using System;
using System.Collections.Generic;

namespace AEIOU
{
    internal delegate string ColumnHeaderReader(int column);
    internal delegate int ColumnUsedCountReader(int column);

    // A complete destination-column snapshot. Keeping metadata and values together
    // preserves the established used-count, header, then cell application order.
    internal sealed class ColumnWriteEntry
    {
        private readonly string header;
        private readonly string[] values;

        public int Column { get; private set; }
        public int UsedCount { get; private set; }

        public string Header
        {
            get { return header ?? String.Empty; }
        }

        public int ValueCount
        {
            get { return values.Length; }
        }

        public string GetValue(int row)
        {
            return values[row] ?? String.Empty;
        }

        public ColumnWriteEntry(int column, int usedCount, string header, string[] values)
        {
            Column = column;
            UsedCount = usedCount;
            this.header = header ?? String.Empty;
            this.values = values ?? new string[0];
        }
    }

    // Calculates column movement without depending on Form1, WinForms, or undo.
    internal static class SheetColumnEditCalculator
    {
        public static IList<ColumnWriteEntry> CreateInsertColumn(
            int rowCount, int originalColumnCount, int column,
            CellValueReader valueReader, ColumnHeaderReader headerReader,
            ColumnUsedCountReader usedCountReader)
        {
            ValidateArguments(rowCount, originalColumnCount, column,
                valueReader, headerReader, usedCountReader);

            List<ColumnWriteEntry> writes = new List<ColumnWriteEntry>();
            for (int sourceColumn = originalColumnCount - 1; sourceColumn >= column; sourceColumn--)
            {
                writes.Add(CreateCopy(sourceColumn, sourceColumn + 1, rowCount,
                    valueReader, headerReader, usedCountReader));
            }
            writes.Add(new ColumnWriteEntry(column, 0, String.Empty, new string[rowCount]));
            return writes;
        }

        public static IList<ColumnWriteEntry> CreateDeleteColumn(
            int rowCount, int columnCount, int column,
            CellValueReader valueReader, ColumnHeaderReader headerReader,
            ColumnUsedCountReader usedCountReader)
        {
            ValidateArguments(rowCount, columnCount, column,
                valueReader, headerReader, usedCountReader);

            List<ColumnWriteEntry> writes = new List<ColumnWriteEntry>();
            for (int sourceColumn = column + 1; sourceColumn < columnCount; sourceColumn++)
            {
                writes.Add(CreateCopy(sourceColumn, sourceColumn - 1, rowCount,
                    valueReader, headerReader, usedCountReader));
            }
            return writes;
        }

        private static ColumnWriteEntry CreateCopy(
            int sourceColumn, int destinationColumn, int rowCount,
            CellValueReader valueReader, ColumnHeaderReader headerReader,
            ColumnUsedCountReader usedCountReader)
        {
            string[] values = new string[rowCount];
            for (int row = 0; row < rowCount; row++)
            {
                values[row] = valueReader(row, sourceColumn) ?? String.Empty;
            }
            return new ColumnWriteEntry(destinationColumn, usedCountReader(sourceColumn),
                headerReader(sourceColumn), values);
        }

        private static void ValidateArguments(
            int rowCount, int columnCount, int column,
            CellValueReader valueReader, ColumnHeaderReader headerReader,
            ColumnUsedCountReader usedCountReader)
        {
            if (rowCount <= 0) throw new ArgumentOutOfRangeException("rowCount");
            if (columnCount <= 0) throw new ArgumentOutOfRangeException("columnCount");
            if (column < 0 || column >= columnCount) throw new ArgumentOutOfRangeException("column");
            if (valueReader == null) throw new ArgumentNullException("valueReader");
            if (headerReader == null) throw new ArgumentNullException("headerReader");
            if (usedCountReader == null) throw new ArgumentNullException("usedCountReader");
        }
    }
}
