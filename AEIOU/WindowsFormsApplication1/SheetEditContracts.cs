using System;
using System.Collections.Generic;

namespace AEIOU
{
    // Calculator input uses row/column order consistently with CellWriteEntry.
    // Implementations are read-only and must not return a mutable cell object.
    internal delegate string CellValueReader(int row, int column);

    // A single, normalized write produced by a sheet-edit calculation.
    internal struct CellWriteEntry
    {
        public readonly int Row;
        public readonly int Col;
        private readonly string value;

        public string Value
        {
            get { return value ?? String.Empty; }
        }

        public CellWriteEntry(int row, int col, string value)
        {
            Row = row;
            Col = col;
            this.value = value ?? String.Empty;
        }
    }

    // Validation belongs to the boundary immediately before host-side application.
    // This keeps a malformed calculation from being applied only partially.
    internal static class CellWriteBatch
    {
        public static void Validate(IList<CellWriteEntry> writes, int rowCount, int columnCount)
        {
            if (writes == null)
            {
                throw new ArgumentNullException("writes");
            }
            if (rowCount <= 0)
            {
                throw new ArgumentOutOfRangeException("rowCount");
            }
            if (columnCount <= 0)
            {
                throw new ArgumentOutOfRangeException("columnCount");
            }

            HashSet<long> coordinates = new HashSet<long>();
            foreach (CellWriteEntry write in writes)
            {
                if (write.Row < 0 || write.Row >= rowCount)
                {
                    throw new ArgumentOutOfRangeException("writes", "A write row is outside the sheet.");
                }
                if (write.Col < 0 || write.Col >= columnCount)
                {
                    throw new ArgumentOutOfRangeException("writes", "A write column is outside the sheet.");
                }

                long coordinate = ((long)write.Row << 32) | (uint)write.Col;
                if (!coordinates.Add(coordinate))
                {
                    throw new ArgumentException("The write list contains a duplicate cell coordinate.", "writes");
                }
            }
        }
    }
}
