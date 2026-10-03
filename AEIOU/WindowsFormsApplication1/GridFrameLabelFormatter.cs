using System;
using System.Collections.Generic;

namespace AEIOU
{
    internal class GridFrameLabelFormatter
    {
        private readonly int fps;
        private readonly int sheetSec;

        public GridFrameLabelFormatter(int fps, int sheetSec)
        {
            this.fps = fps;
            this.sheetSec = sheetSec;
        }

        // フレーム表示文字列の生成
        public string FrmToSheet(int frm, IList<Range> addRange)
        {
            string Time = "";
            int addFrame = 0;
            int localFrm = 0;
            bool isAddRange = false;

            //切り貼りフレームのカウント
            foreach (Range r in addRange)
            {
                //切り貼りコマ数の積算
                if (r.Top <= frm)
                {
                    addFrame += (r.Bottom - r.Top + 1);
                }

                //切り貼り範囲のチェック
                if (r.Top <= frm && r.Bottom >= frm)
                {
                    //範囲内だったら、切り貼り内ローカルのコマ数を計算
                    // ※また、フレーム表示は１～ なので更に＋１
                    isAddRange = true;
                    localFrm = (frm - r.Top) + 1;
                }
            }

            // カレントフレームに、切り貼り分を加味
            // ※また、フレーム表示は１～ なので更に＋１
            frm = (frm - addFrame) + 1;

            // 24fps
            if (this.fps == 24)
            {
                if (isAddRange)
                {
                    // 付けたしフレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = 0;  //ページ数は0
                    int frm_f = (localFrm - 1) % (this.sheetSec * 24);

                    if (((localFrm - 1) % 12) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();

                }
                else
                {
                    //通常フレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = (frm - 1) / (this.sheetSec * this.fps) + 1;    //ページ数は１から
                    int frm_f = (frm - 1) % (this.sheetSec * this.fps);

                    if (((frm - 1) % 12) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();
                }
            }

            // 1 & 30fps
            if ((this.fps == 1) || (this.fps == 30))
            {
                if (isAddRange)
                {
                    // 付けたしフレーム表示処理
                    // フレーム計算(Page + f)
                    int frm_p = 0;  //ページ数は0
                    int frm_f = (localFrm - 1) % (this.sheetSec * this.fps);

                    if (((localFrm - 1) % 15) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();

                }
                else
                {
                    //通常フレーム表示処理

                    // フレーム計算(Page + f)
                    int frm_p = (frm - 1) / (this.sheetSec * this.fps) + 1;    //ページ数は１から
                    int frm_f = (frm - 1) % (this.sheetSec * this.fps);

                    if (((frm - 1) % 15) == 0)
                    {
                        if (frm_p < 10) Time = "0" + frm_p.ToString() + '/';
                        else Time = frm_p.ToString() + '/';
                    }
                    else
                    {
                        Time = "     ";
                    }

                    if ((frm_f + 1) < 10) Time += "0" + (frm_f + 1).ToString();
                    else Time += (frm_f + 1).ToString();
                }
            }

            return Time;
        }
    }
}
