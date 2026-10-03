using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace AEIOU
{
    internal class GridFrameHeaderPainter
    {
        private readonly GridFrameLabelFormatter formatter;

        public GridFrameHeaderPainter(GridFrameLabelFormatter formatter)
        {
            this.formatter = formatter;
        }

        public void DrawHeader(
            DataGridViewCellPaintingEventArgs e,
            Color headerColor,
            bool isDisplayFrameNumber,
            int firstFrame,
            IList<Range> addRange)
        {
            // 背景塗り
            using (Brush backColorBrush = new SolidBrush(headerColor))
            {
                e.Graphics.FillRectangle(backColorBrush, e.CellBounds);
            }

            //範囲を取得
            Rectangle _rect = e.CellBounds;
            _rect.Inflate(-2, -2);
            //文字列を描画
            String value = "";
            if(isDisplayFrameNumber)
            {
                // フレーム数 表示
                value = (e.RowIndex + firstFrame).ToString();
            }
            else
            {
                // シート/コマ数 表示
                value = formatter.FrmToSheet(e.RowIndex, addRange);
            }
            TextRenderer.DrawText(
                e.Graphics,
                value,
                e.CellStyle.Font,
                _rect,
                e.CellStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);

            //背景(及びヘッダー部の三角カーソル)以外を描画して貰う
            DataGridViewPaintParts _paintParts =
                e.PaintParts & ~DataGridViewPaintParts.Background;
            //残りの描画を要求
            e.Paint(e.ClipBounds, _paintParts);

            //描画完了の通知
            e.Handled = true;
        }
    }
}
