using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    // デフォルトのColumnクラスを継承し TimingColumnクラスを定義
    class TimingColumn : DataGridViewColumn
    {
        public TimingColumn()
            : base(new TimingCell())
        {
        }

        // CellTemplate を override し、CalendarCell のみ受け入れるように制限する
        // ※CalendarCell 以外を設定された場合には例外を throw する
        public override DataGridViewCell CellTemplate
        {
            get
            {
                return base.CellTemplate;
            }
            set
            {
                if (value != null &&
                    !value.GetType().IsAssignableFrom(typeof(TimingCell)))
                {
                    throw new InvalidCastException("TimingCellのみ入力可能.");
                }
                base.CellTemplate = value;
            }
        }
   }
}
