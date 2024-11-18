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

        // ここに DataGridViewの実装を書くのは変か..
        public void CopyToBuffer(Rect range, GridViewManager manager)
        {
            manager.CopyRect = range;
            manager.CopyBuffer = new string[range.Height, range.Width];
            for (int i = 0; i < range.Height; i++)
            {
                for (int j = 0; j < range.Width; j++)
                {
                    manager.CopyBuffer[i, j] = manager.View[range.X + j, range.Y + i].Value.ToString();
                }
            }
        }
        public void PasteFromBuffer(int row, int col, GridViewManager manager)
        {
            if (manager.CopyBuffer != null)
            {
                Rect copyRect = manager.CopyRect;
                for (int i = 0; i < Math.Min(copyRect.Height, manager.CopyBuffer.GetLength(0)); i++)
                {
                    for (int j = 0; j < Math.Min(copyRect.Width, manager.CopyBuffer.GetLength(1)); j++)
                    {
                        manager.View[col + j, row + i].Value = manager.CopyBuffer[i, j];
                    }
                }
            }
        }
        public void ClearSelection(Rect range, GridViewManager manager)
        {
            for (int i = 0; i < range.Height; i++)
            {
                for (int j = 0; j < range.Width; j++)
                {
                    manager.View[range.X + j, range.Y + i].Value = "";
                }
            }
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
            _oldValue = manager.View[_column, _row].Value.ToString();
            manager.View[_column, _row].Value = _newValue;
        }

        public override void Undo(GridViewManager manager)
        {
            manager.View[_column, _row].Value = _oldValue;
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
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

        public PasteOperation(int row, int col)
        {
            _name = "貼り付け";
            _row = row;
            _col = col;
        }

        public override void Execute(GridViewManager manager)
        {
            if (manager.CopyBuffer != null)
            {
                Rect copyRect = manager.CopyRect;
                _oldValues = new String[copyRect.Height, copyRect.Width];
                for (int i = 0; i < copyRect.Height; i++)
                {
                    for (int j = 0; j < copyRect.Width; j++)
                    {
                        _oldValues[i, j] = manager.View[_col + j, _row + i].Value.ToString();
                    }
                }
                PasteFromBuffer(_row, _col, manager);
            }
        }

        public override void Undo(GridViewManager manager)
        {
            if (manager.CopyBuffer != null)
            {
                int rowCount = _oldValues.GetLength(0);
                int columnCount = _oldValues.GetLength(1);
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        manager.View[_col + j, _row + i].Value = _oldValues[i, j];
                    }
                }
            }
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
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
            _oldValues = new String[_cutRange.Height, _cutRange.Width];
            for (int i = 0; i < _cutRange.Height; i++)
            {
                for (int j = 0; j < _cutRange.Width; j++)
                {
                    _oldValues[i, j] = manager.View[_cutRange.X + j, _cutRange.Y + i].Value.ToString();
                }
            }
            CopyToBuffer(_cutRange, manager);
            ClearSelection(_cutRange, manager);
        }

        public override void Undo(GridViewManager manager)
        {
            for (int i = 0; i < _cutRange.Height; i++)
            {
                for (int j = 0; j < _cutRange.Width; j++)
                {
                    manager.View[_cutRange.X + j, _cutRange.Y + i].Value = _oldValues[i, j];
                }
            }
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
            _oldValues = new String[_deleteRange.Height, _deleteRange.Width];
            for (int i = 0; i < _deleteRange.Height; i++)
            {
                for (int j = 0; j < _deleteRange.Width; j++)
                {
                    _oldValues[i, j] = manager.View[_deleteRange.X + j, _deleteRange.Y + i].Value.ToString();
                }
            }
            ClearSelection(_deleteRange, manager);
        }

        public override void Undo(GridViewManager manager)
        {
            for (int i = 0; i < _deleteRange.Height; i++)
            {
                for (int j = 0; j < _deleteRange.Width; j++)
                {
                    manager.View[_deleteRange.X + j, _deleteRange.Y + i].Value = _oldValues[i, j];
                }
            }
        }

        public override void Redo(GridViewManager manager)
        {
            Execute(manager);
        }

    }

}
