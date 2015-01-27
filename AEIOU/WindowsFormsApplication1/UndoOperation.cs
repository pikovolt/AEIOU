using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AEIOU
{
    // UNDO用 操作内容保持
    class UndoOperation
    {
        private int row;            // セルの行数
        private int col;            // 　　　列数
        private string oldValue;    // 編集前の値
        private string newValue;    // 編集後の値

        public UndoOperation(
            int _col, int _row,
            string _old, string _new)
        {
            this.col = _col;
            this.row = _row;
            this.oldValue = _old;
            this.newValue = _new;
        }

        public int Row
        {
            get
            {
                return this.row;
            }
            set
            {
                this.row = value;
            }
        }

        public int Column
        {
            get
            {
                return this.col;
            }
            set
            {
                this.col = value;
            }
        }

        public String NewValue
        {
            get
            {
                return this.newValue;
            }
            set
            {
                this.newValue = value;
            }
        }

        public String OldValue
        {
            get
            {
                return this.oldValue;
            }
            set
            {
                this.oldValue = value;
            }
        }
    }
}
