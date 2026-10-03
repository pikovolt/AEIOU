using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    public class OperationGroup : GridViewOperation
    {
        private List<GridViewOperation> _operations = new List<GridViewOperation>();

        public OperationGroup(String name)
        {
            _name = name;
        }

        public void AddOperation(GridViewOperation operation)
        {
            _operations.Add(operation);
        }

        public override void Execute(GridViewManager manager)
        {
            foreach (var operation in _operations)
            {
                operation.Execute(manager);
            }
        }

        public override void Undo(GridViewManager manager)
        {
            for (int i = _operations.Count - 1; i >= 0; i--)
            {
                _operations[i].Undo(manager);
            }
        }

        public override void Redo(GridViewManager manager)
        {
            foreach (var operation in _operations)
            {
                operation.Redo(manager);
            }
        }

        // _operationsの数を取得するためのプロパティを追加
        public int OperationCount => _operations.Count;

    }
}
