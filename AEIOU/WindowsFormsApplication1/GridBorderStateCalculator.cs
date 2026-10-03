using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace AEIOU
{
    internal class GridBorderStateCalculator
    {
        private readonly int fps;
        private readonly int sheetSec;
        private readonly int sheetDivide;

        public GridBorderStateCalculator(int fps, int sheetSec, int sheetDivide)
        {
            this.fps = fps;
            this.sheetSec = sheetSec;
            this.sheetDivide = sheetDivide;
        }

        public SheetBorder CalcBorderState(
            DataGridViewCellPaintingEventArgs e,
            IList<Range> addRange)
        {
            SheetBorder retValue = SheetBorder.None;
            bool bHariFirst = false;
            int addFrame = 0;

            // 切り貼りフレーム数のカウント
            {
                // 先頭フレーム前の切り貼り有無
                if(addRange.Count > 0)
                {
                    if(addRange[0].Top == 0)
                    {
                        //有り
                        bHariFirst = true;
                    }
                }

                // 追加のコマ数を数える
                foreach(Range range in addRange)
                {
                    if(range.Top <= e.RowIndex)
                    {
                        addFrame += (range.Bottom - range.Top + 1);
                    }
                }
            }

            // [12],24フレーム毎に基準線を引く
            // (１シート毎に基準線を引く：ページ線)
            int r = e.RowIndex - addFrame + 1;
            if (r < 1) r = e.RowIndex + 1;
            if (this.fps == 24 && e.ColumnIndex >= 0)
            {
                if ((((r % this.sheetDivide) == 0) && (r != 0)) || ((r == addFrame) && ((e.RowIndex + 1) - addFrame == 0) && (addFrame != 0) && (r != 0)))
                {
                    if ((r % 24) == 0)
                    {
                        //１秒毎の基準線を描画
                        retValue = SheetBorder.EverySec;
                    }
                    else
                    {
                        //ｎコマ毎の基準線を描画
                        retValue = SheetBorder.EveryNFrames;
                    }

                    if ((r % (this.sheetSec * this.fps)) == 0 || ((e.RowIndex + 1) == addFrame) && bHariFirst == true)
                    {
                        //シート毎の基準線を描画
                        retValue = SheetBorder.EverySheet;
                    }
                }
            }

            // 1 & 30fps
            if (((this.fps == 1) || (this.fps == 30)) && e.ColumnIndex >= 0)
            {
                // 15,30フレーム毎に基準線を引く
                // (１シート毎に基準線を引く：ページ線)
                if ((((r % 15) == 0) && (r != 0)) || ((r == addFrame) && ((e.RowIndex + 1) - addFrame == 0) && (addFrame != 0) && (r != 0)))
                {
                    if ((r % 30) == 0)
                    {
                        //１秒毎の基準線を描画
                        retValue = SheetBorder.EverySec;
                    }
                    else
                    {
                        //ｎコマ毎の基準線を描画
                        retValue = SheetBorder.EveryNFrames;
                    }

                    if ((r % (this.sheetSec * this.fps)) == 0 || ((e.RowIndex + 1) == addFrame) && bHariFirst == true)
                    {
                        //シート毎の基準線を描画
                        retValue = SheetBorder.EverySheet;
                    }
                }
            }
            return retValue;
        }
    }
}
