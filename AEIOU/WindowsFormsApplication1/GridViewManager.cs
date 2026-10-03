using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    public class GridViewManager
    {
        public delegate void CellValueChangedHandler(int col, int row, string value);

        private UndoManager _undoManager;
        private DataGridView _view;
        private string[,] _copyBuffer;              // DataGridView向けのコピーバッファを保持
        private Rect _copyRect;                     // コピー範囲を保持
        private Stack<OperationGroup> _groupStack;  // OperationGroupの入れ子対応
        private TimingSheetModel _model;
        private int _batchUpdateDepth;
        private Dictionary<int, CellChange> _pendingCellChanges;

        private struct CellChange
        {
            public readonly int Row;
            public readonly string Value;

            public CellChange(int row, string value)
            {
                Row = row;
                Value = value;
            }
        }

        public event CellValueChangedHandler CellValueChanged;

        public DataGridView View
        {
            get { return _view; }
            set { _view = value; }
        }

        public string[,] CopyBuffer
        {
            get { return _copyBuffer; }
            set { _copyBuffer = value; }
        }
        public Rect CopyRect
        {
            get { return _copyRect; }
            set { _copyRect = value; }
        }

        public TimingSheetModel Model
        {
            get { return _model; }
            set { _model = value; }
        }

        public int ColumnCount
        {
            get
            {
                if (_model != null)
                {
                    return _model.ColumnCount;
                }

                return (_view != null) ? _view.ColumnCount : 0;
            }
        }

        public int RowCount
        {
            get
            {
                if (_model != null)
                {
                    return _model.RowCount;
                }

                return (_view != null) ? _view.RowCount : 0;
            }
        }

        public void InitializeWork(DataGridView view, TimingSheetModel model)
        {
            _view = view;
            _model = model;
            _copyBuffer = null;
            _undoManager = new UndoManager();
            _groupStack = new Stack<OperationGroup>();
            _batchUpdateDepth = 0;
            _pendingCellChanges = new Dictionary<int, CellChange>();
        }

        public OperationGroup BeginGroup(String name)
        {
            var newGroup = new OperationGroup(name);
            _groupStack.Push(newGroup);
            return newGroup;
        }

        public void EndGroup()
        {
            if (_groupStack.Count > 0)
            {
                var completedGroup = _groupStack.Pop();

                if (_groupStack.Count == 0 && completedGroup.OperationCount > 0)
                {
                    _undoManager.PushOperation(completedGroup);
                }
                else if (_groupStack.Count > 0)
                {
                    _groupStack.Peek().AddOperation(completedGroup);
                }
            }
        }
        
        public void ExecuteOperation(GridViewOperation operation)
        {
            if (_groupStack.Count > 0)
            {
                _groupStack.Peek().AddOperation(operation);
            }
            else
            {
                _undoManager.PushOperation(operation);
            }
            operation.Execute(this);
        }

        public void Undo()
        {
            BeginBatchUpdate();
            try
            {
                _undoManager.Undo(this);
            }
            finally
            {
                EndBatchUpdate();
            }
        }

        public bool TryUndo(GridViewOperation expectedOperation)
        {
            BeginBatchUpdate();
            try
            {
                return _undoManager.TryUndo(this, expectedOperation);
            }
            finally
            {
                EndBatchUpdate();
            }
        }

        public void Redo()
        {
            BeginBatchUpdate();
            try
            {
                _undoManager.Redo(this);
            }
            finally
            {
                EndBatchUpdate();
            }
        }

        public void BeginBatchUpdate()
        {
            _batchUpdateDepth++;
        }

        public void EndBatchUpdate()
        {
            if (_batchUpdateDepth <= 0)
            {
                throw new InvalidOperationException("EndBatchUpdate requires a matching BeginBatchUpdate.");
            }

            _batchUpdateDepth--;
            if (_batchUpdateDepth != 0 || _pendingCellChanges.Count == 0)
            {
                return;
            }

            // A row shift changes many cells in the same column. Notify once per changed
            // column so continuity state is recalculated only once for that column.
            var pendingChanges = new Dictionary<int, CellChange>(_pendingCellChanges);
            _pendingCellChanges.Clear();
            foreach (KeyValuePair<int, CellChange> entry in pendingChanges)
            {
                NotifyCellValueChanged(entry.Key, entry.Value.Row, entry.Value.Value);
            }
        }

        public string GetCellValue(int col, int row)
        {
            return _model.GetCell(col, row);
        }

        public void SyncGridShape(int columnCount, int rowCount)
        {
            if (_view == null)
            {
                return;
            }

            _view.RowCount = rowCount;
        }

        public bool TryHandleCellValueNeeded(int col, int row, out string value)
        {
            return TryGetCellValue(col, row, out value);
        }

        public void PushCellValue(int col, int row, object rawValue)
        {
            SetCellValue(col, row, rawValue == null ? string.Empty : rawValue.ToString());
        }

        public bool TryGetCellValue(int col, int row, out string value, out string failureReason)
        {
            if (_model == null)
            {
                value = string.Empty;
                failureReason = "ModelNotInitialized";
                return false;
            }

            return _model.TryGetCell(col, row, out value, out failureReason);
        }

        public bool TryGetCellValue(int col, int row, out string value)
        {
            string failureReason;
            return TryGetCellValue(col, row, out value, out failureReason);
        }

        public void SetCellValue(int col, int row, string value)
        {
            EnsureModelBoundForWrite();
            string normalizedValue = NormalizeCellValue(value);
            _model.SetCell(col, row, normalizedValue);
            SetCellDisplayValue(col, row, normalizedValue);
            NotifyCellValueChanged(col, row, normalizedValue);
        }

        public string SetCellValueWithUndo(int col, int row, string value)
        {
            EnsureModelBoundForWrite();
            string normalizedValue = NormalizeCellValue(value);
            string oldValue = _model.SetCellWithUndo(col, row, normalizedValue);
            SetCellDisplayValue(col, row, normalizedValue);
            NotifyCellValueChanged(col, row, normalizedValue);
            return oldValue;
        }

        public void ApplyUndoCellValue(int col, int row, string value)
        {
            EnsureModelBoundForWrite();
            string normalizedValue = NormalizeCellValue(value);
            _model.ApplyUndoCell(col, row, normalizedValue);
            SetCellDisplayValue(col, row, normalizedValue);
            NotifyCellValueChanged(col, row, normalizedValue);
        }

        public void ApplyRedoCellValue(int col, int row, string value)
        {
            EnsureModelBoundForWrite();
            string normalizedValue = NormalizeCellValue(value);
            _model.ApplyRedoCell(col, row, normalizedValue);
            SetCellDisplayValue(col, row, normalizedValue);
            NotifyCellValueChanged(col, row, normalizedValue);
        }

        public string GetHeaderValue(int col)
        {
            return _model.GetHeader(col);
        }

        public void SetHeaderValue(int col, string value)
        {
            EnsureModelBoundForWrite();
            string normalizedValue = NormalizeCellValue(value);
            _model.SetHeader(col, normalizedValue);
            SetHeaderDisplayValue(col, normalizedValue);
        }

        public void SetCellDisplayValue(int col, int row, string value)
        {
            if (_view == null)
            {
                return;
            }

            if (_view.VirtualMode)
            {
                if (_batchUpdateDepth > 0)
                {
                    return;
                }

                if (col >= 0 && row >= 0 && col < _view.ColumnCount && row < _view.RowCount)
                {
                    _view.InvalidateCell(col, row);
                }

                return;
            }

            _view[col, row].Value = NormalizeCellValue(value);
        }

        public void SetHeaderDisplayValue(int col, string value)
        {
            if (_view == null)
            {
                return;
            }

            _view.Columns[col].HeaderText = NormalizeCellValue(value);
        }

        public string[,] GetRangeValues(Rect range)
        {
            ValidateRange(range);
            string[,] values = new string[range.Height, range.Width];
            for (int rowOffset = 0; rowOffset < range.Height; rowOffset++)
            {
                for (int colOffset = 0; colOffset < range.Width; colOffset++)
                {
                    values[rowOffset, colOffset] = GetCellValue(range.X + colOffset, range.Y + rowOffset);
                }
            }

            return values;
        }

        public void SetRangeValues(int startCol, int startRow, string[,] values)
        {
            if (values == null)
            {
                throw new ArgumentNullException("values");
            }

            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);
            ValidateRange(new Rect(startCol, startRow, columnCount, rowCount));
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                for (int colOffset = 0; colOffset < columnCount; colOffset++)
                {
                    SetCellValue(startCol + colOffset, startRow + rowOffset, values[rowOffset, colOffset]);
                }
            }
        }

        public void ClearRangeValues(Rect range)
        {
            ValidateRange(range);
            for (int rowOffset = 0; rowOffset < range.Height; rowOffset++)
            {
                for (int colOffset = 0; colOffset < range.Width; colOffset++)
                {
                    SetCellValue(range.X + colOffset, range.Y + rowOffset, "");
                }
            }
        }


        private void ValidateRange(Rect range)
        {
            if (range.Width < 0 || range.Height < 0)
            {
                throw new ArgumentOutOfRangeException("range", "Range size must be non-negative.");
            }

            if (range.Width == 0 || range.Height == 0)
            {
                return;
            }

            ValidateCellIndex(range.X, range.Y);
            ValidateCellIndex(range.Right, range.Bottom);
        }

        private void ValidateCellIndex(int col, int row)
        {
            if (col < 0 || row < 0 || col >= ColumnCount || row >= RowCount)
            {
                throw new ArgumentOutOfRangeException("col,row", "Cell index is out of range.");
            }
        }

        private string NormalizeCellValue(string value)
        {
            return value ?? "";
        }

        private void EnsureModelBoundForWrite()
        {
            if (_model == null)
            {
                throw new InvalidOperationException("GridViewManager write operation requires a bound model. Call InitializeWork(view, model) before write APIs.");
            }
        }

        private void NotifyCellValueChanged(int col, int row, string value)
        {
            if (_batchUpdateDepth > 0)
            {
                _pendingCellChanges[col] = new CellChange(row, value);
                return;
            }

            CellValueChangedHandler handler = CellValueChanged;
            if (handler != null)
            {
                handler(col, row, value);
            }
        }

    }
}
