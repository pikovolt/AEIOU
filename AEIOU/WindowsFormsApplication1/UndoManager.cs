using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    internal class UndoManager
    {
        private List<GridViewOperation> _undoStack;
        private List<GridViewOperation> _redoStack;

        public UndoManager()
        {
            _undoStack = new List<GridViewOperation>();
            _redoStack = new List<GridViewOperation>();
        }

        public void PushOperation(GridViewOperation operation)
        {
            _undoStack.Add(operation);
            _redoStack.Clear();
        }

        public void Undo(GridViewManager manager)
        {
            if (_undoStack.Count > 0)
            {
                var operation = _undoStack[_undoStack.Count - 1];
                _undoStack.RemoveAt(_undoStack.Count - 1);
                operation.Undo(manager);
                _redoStack.Add(operation);
            }
        }

        public void Redo(GridViewManager manager)
        {
            if (_redoStack.Count > 0)
            {
                var operation = _redoStack[_redoStack.Count - 1];
                _redoStack.RemoveAt(_redoStack.Count - 1);
                operation.Redo(manager);
                _undoStack.Add(operation);
            }
        }
    }
}
