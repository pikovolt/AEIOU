using System;
using System.Collections.Generic;

namespace AEIOU
{
    // Calculates row-edit writes without depending on the UI, model, or undo implementation.
    internal static class SheetRowEditCalculator
    {
        public static IList<CellWriteEntry> CreateInsertRows(
            int rowCount, int columnCount, int row, int count, CellValueReader valueReader)
        {
            ValidateArguments(rowCount, columnCount, row, count, valueReader);

            List<CellWriteEntry> writes = new List<CellWriteEntry>();
            int movableLength = rowCount - row - count;
            for (int column = 0; column < columnCount; column++)
            {
                // Preserve the established order: a downward shift is emitted from the end.
                for (int offset = movableLength - 1; offset >= 0; offset--)
                {
                    int sourceRow = row + offset;
                    writes.Add(new CellWriteEntry(
                        sourceRow + count, column, valueReader(sourceRow, column)));
                }
            }

            AddClearWrites(writes, columnCount, row, count);
            return writes;
        }

        public static IList<CellWriteEntry> CreateDeleteRows(
            int rowCount, int columnCount, int row, int count, CellValueReader valueReader)
        {
            ValidateArguments(rowCount, columnCount, row, count, valueReader);

            List<CellWriteEntry> writes = new List<CellWriteEntry>();
            int sourceStartRow = row + count;
            int movableLength = rowCount - sourceStartRow;
            for (int column = 0; column < columnCount; column++)
            {
                // Preserve the established order: an upward shift is emitted from the start.
                for (int offset = 0; offset < movableLength; offset++)
                {
                    int sourceRow = sourceStartRow + offset;
                    writes.Add(new CellWriteEntry(
                        row + offset, column, valueReader(sourceRow, column)));
                }
            }

            AddClearWrites(writes, columnCount, rowCount - count, count);
            return writes;
        }

        private static void AddClearWrites(
            IList<CellWriteEntry> writes, int columnCount, int startRow, int count)
        {
            // ClearRangeValues historically visits a rectangular range row first, then column.
            for (int rowOffset = 0; rowOffset < count; rowOffset++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    writes.Add(new CellWriteEntry(startRow + rowOffset, column, String.Empty));
                }
            }
        }

        private static void ValidateArguments(
            int rowCount, int columnCount, int row, int count, CellValueReader valueReader)
        {
            if (rowCount <= 0)
            {
                throw new ArgumentOutOfRangeException("rowCount");
            }
            if (columnCount <= 0)
            {
                throw new ArgumentOutOfRangeException("columnCount");
            }
            if (row < 0 || row >= rowCount)
            {
                throw new ArgumentOutOfRangeException("row");
            }
            if (count <= 0 || count > rowCount - row)
            {
                throw new ArgumentOutOfRangeException("count");
            }
            if (valueReader == null)
            {
                throw new ArgumentNullException("valueReader");
            }
        }
    }
}
