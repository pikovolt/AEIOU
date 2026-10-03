using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    public abstract class GridViewOperation
    {
        protected String _name;

        public String Name
        {
            get { return _name; }
        }

        public abstract void Execute(GridViewManager manager);
        public abstract void Undo(GridViewManager manager);
        public abstract void Redo(GridViewManager manager);

        public void CopyToBuffer(Rect range, GridViewManager manager)
        {
            manager.CopyRect = range;
            manager.CopyBuffer = manager.GetRangeValues(range);
        }

        public void PasteFromBuffer(int row, int col, GridViewManager manager)
        {
            if (manager.CopyBuffer != null)
            {
                // コピー範囲のサイズを取得
                int bufferHeight = manager.CopyBuffer.GetLength(0);
                int bufferWidth = manager.CopyBuffer.GetLength(1);

                // DataGridViewの範囲内に収まるように調整
                int maxHeight = Math.Min(bufferHeight, manager.RowCount - row);
                int maxWidth = Math.Min(bufferWidth, manager.ColumnCount - col);
                if (maxHeight <= 0 || maxWidth <= 0)
                {
                    return;
                }

                string[,] clippedValues = new string[maxHeight, maxWidth];
                for (int i = 0; i < maxHeight; i++)
                {
                    for (int j = 0; j < maxWidth; j++)
                    {
                        clippedValues[i, j] = manager.CopyBuffer[i, j];
                    }
                }

                manager.SetRangeValues(col, row, clippedValues);
            }
        }


        public void ClearSelection(Rect range, GridViewManager manager)
        {
            manager.ClearRangeValues(range);
        }

    }
    public class SetValueOperation : GridViewOperation
    {
        private int _row;
        private int _column;
        private string _newValue;
        private string _oldValue;

        public SetValueOperation(int row, int column, string newValue)
        {
            _name = "入力";
            _row = row;
            _column = column;
            _newValue = newValue;
        }

        public override void Execute(GridViewManager manager)
        {
            _oldValue = manager.SetCellValueWithUndo(_column, _row, _newValue);
        }

        public override void Undo(GridViewManager manager)
        {
            manager.ApplyUndoCellValue(_column, _row, _oldValue);
        }

        public override void Redo(GridViewManager manager)
        {
            manager.ApplyRedoCellValue(_column, _row, _newValue);
        }
    }

    public class CopyOperation : GridViewOperation
    {
        private Rect _copyRange;

        public CopyOperation(Rect copyRange)
        {
            _name = "複製";
            _copyRange = copyRange;
        }

        public override void Execute(GridViewManager manager)
        {
            CopyToBuffer(_copyRange, manager);
        }

        public override void Undo(GridViewManager manager)
        {
            // (コピーの undo時は、なにもしない)
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
        }
    }

    public class PasteOperation : GridViewOperation
    {
        private int _col;
        private int _row;
        private String[,] _oldValues;
        private String[,] _newValues;

        public PasteOperation(int row, int col)
        {
            _name = "貼り付け";
            _row = row;
            _col = col;
        }

        public override void Execute(GridViewManager manager)
        {
            if (manager.CopyBuffer == null)
            {
                _oldValues = null;
                _newValues = null;
                return;
            }

            // 範囲外の処理はしないよう、コピー範囲を計算
            Rect copyRect = manager.CopyRect;
            int maxHeight = Math.Min(copyRect.Height, manager.RowCount - _row);
            int maxWidth = Math.Min(copyRect.Width, manager.ColumnCount - _col);
            if (maxHeight <= 0 || maxWidth <= 0)
            {
                _oldValues = null;
                _newValues = null;
                return;
            }

            Rect pasteRange = new Rect(_col, _row, maxWidth, maxHeight);

            _oldValues = manager.GetRangeValues(pasteRange);
            _newValues = new String[maxHeight, maxWidth];
            for (int i = 0; i < maxHeight; i++)
            {
                for (int j = 0; j < maxWidth; j++)
                {
                    _newValues[i, j] = manager.CopyBuffer[i, j];
                }
            }

            manager.SetRangeValues(_col, _row, _newValues);
        }

        public override void Undo(GridViewManager manager)
        {
            if (_oldValues == null)
            {
                return;
            }

            manager.SetRangeValues(_col, _row, _oldValues);
        }

        public override void Redo(GridViewManager manager)
        {
            if (_newValues == null)
            {
                return;
            }

            manager.SetRangeValues(_col, _row, _newValues);
        }
    }

    public class CutOperation : GridViewOperation
    {
        private Rect _cutRange;
        private String[,] _oldValues;

        public CutOperation(Rect cutRange)
        {
            _name = "切り取り";
            _cutRange = cutRange;
        }

        public override void Execute(GridViewManager manager)
        {
            _oldValues = manager.GetRangeValues(_cutRange);
            CopyToBuffer(_cutRange, manager);
            ClearSelection(_cutRange, manager);
        }

        public override void Undo(GridViewManager manager)
        {
            manager.SetRangeValues(_cutRange.X, _cutRange.Y, _oldValues);
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
        }

    }

    public class DeleteOperation : GridViewOperation
    {
        private Rect _deleteRange;
        private String[,] _oldValues;

        public DeleteOperation(Rect deleteRange)
        {
            _name = "削除";
            _deleteRange = deleteRange;
        }

        public override void Execute(GridViewManager manager)
        {
            _oldValues = manager.GetRangeValues(_deleteRange);
            ClearSelection(_deleteRange, manager);
        }

        public override void Undo(GridViewManager manager)
        {
            manager.SetRangeValues(_deleteRange.X, _deleteRange.Y, _oldValues);
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
        }

    }

}
