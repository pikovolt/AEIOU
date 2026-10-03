using System;

namespace AEIOU
{
    public class TimingSheetModel
    {
        public const string TryGetCellFailureReasonNone = "";
        public const string TryGetCellFailureReasonOutOfRange = "OutOfRange";

        private readonly string[,] _cells;
        private readonly string[] _headers;

        public TimingSheetModel(int columnCount, int rowCount)
        {
            if (columnCount < 0) throw new ArgumentOutOfRangeException("columnCount");
            if (rowCount < 0) throw new ArgumentOutOfRangeException("rowCount");

            _cells = new string[columnCount, rowCount];
            _headers = new string[columnCount];

            for (int col = 0; col < columnCount; col++)
            {
                _headers[col] = "";
                for (int row = 0; row < rowCount; row++)
                {
                    _cells[col, row] = "";
                }
            }
        }

        public int ColumnCount
        {
            get { return _cells.GetLength(0); }
        }

        public int RowCount
        {
            get { return _cells.GetLength(1); }
        }

        public string GetCell(int col, int row)
        {
            ValidateCellIndex(col, row);
            return _cells[col, row];
        }

        public bool TryGetCell(int col, int row, out string value, out string failureReason)
        {
            if (!IsInRange(col, row))
            {
                value = string.Empty;
                failureReason = TryGetCellFailureReasonOutOfRange;
                return false;
            }

            value = _cells[col, row];
            failureReason = TryGetCellFailureReasonNone;
            return true;
        }

        public bool TryGetCell(int col, int row, out string value)
        {
            string failureReason;
            return TryGetCell(col, row, out value, out failureReason);
        }

        public void SetCell(int col, int row, string value)
        {
            ValidateCellIndex(col, row);
            _cells[col, row] = value ?? "";
        }

        public string SetCellWithUndo(int col, int row, string newValue)
        {
            ValidateCellIndex(col, row);
            string oldValue = _cells[col, row];
            _cells[col, row] = newValue ?? "";
            return oldValue;
        }

        public void ApplyUndoCell(int col, int row, string oldValue)
        {
            SetCell(col, row, oldValue);
        }

        public void ApplyRedoCell(int col, int row, string newValue)
        {
            SetCell(col, row, newValue);
        }

        public string GetHeader(int col)
        {
            ValidateColumnIndex(col);
            return _headers[col];
        }

        public void SetHeader(int col, string header)
        {
            ValidateColumnIndex(col);
            _headers[col] = header ?? "";
        }

        private void ValidateCellIndex(int col, int row)
        {
            if (!IsInRange(col, row))
            {
                throw new ArgumentOutOfRangeException("col,row", "Cell index is out of range.");
            }
        }

        private void ValidateColumnIndex(int col)
        {
            if (col < 0 || col >= ColumnCount)
            {
                throw new ArgumentOutOfRangeException("col", "Column index is out of range.");
            }
        }

        private bool IsInRange(int col, int row)
        {
            return col >= 0 && row >= 0 && col < ColumnCount && row < RowCount;
        }
    }
}
