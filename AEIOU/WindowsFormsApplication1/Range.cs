using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AEIOU
{
    // 範囲記録用クラス
    class Range
    {
        private int top;
        private int bottom;

        public Range(int t, int b)
        {
            this.top = t;
            this.bottom = b;
        }

        public int Top
        {
            get
            {
                return this.top;
            }
            set
            {
                this.top = value;
            }
        }

        public int Bottom
        {
            get
            {
                return this.bottom;
            }
            set
            {
                this.bottom = value;
            }
        }

        public int Length
        {
            get
            {
                return this.bottom - this.top + 1;
            }
        }
    }

}
