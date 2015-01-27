using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace AEIOU
{
    // UNDO用 グループ
    class UndoGroup
    {
        public int Num;                         // アンドゥ番号(回数カウントのコピー)
        public string Title;                    // 操作(内容説明用文字列)
        public List<UndoOperation> Operations;  // 操作内容
        public UndoOperationTarget Target;      // 編集対象
        public Rect Rect;                       // 操作範囲(ペースト時のコピーバッファ範囲など)

        public UndoGroup(
            int _num, String _title, UndoOperationTarget _target)
        {
            this.Num = _num;
            this.Title = _title;
            this.Target = _target;
            Operations = new List<UndoOperation>();
            Rect = new Rect();
        }
    }
}
