using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    public class GridViewManager
    {
        private UndoManager _undoManager;
        private DataGridView _view;
        private string[,] _copyBuffer;              // DataGridView向けのコピーバッファを保持
        private Rect _copyRect;                     // コピー範囲を保持
        private Stack<OperationGroup> _groupStack;  // OperationGroupの入れ子対応

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

        public void InitializeWork(DataGridView view)
        {
            _view = view;
            _copyBuffer = null;
            _undoManager = new UndoManager();
            _groupStack = new Stack<OperationGroup>();
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
            _undoManager.Undo(this);
        }

        public void Redo()
        {
            _undoManager.Redo(this);
        }

    }
}
