using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AEIOU
{
    // 文字表示の実装を持つ TextBoxCellクラスを継承し CalendarCellクラスを定義
    class TimingCell : DataGridViewTextBoxCell
    {
        public TimingCell()
            : base()
        {
        }

        // カスタム プロパティ実装
        private bool isHeader = false;
        private bool isKaraCell = false;
        private bool isContinuousLine = false;
        private SheetBorder borderState = SheetBorder.None;
        public bool IsHeader
        {
            set
            {
                isHeader = value;
            }
            get
            {
                return isHeader;
            }
        }
        public bool IsKaraCell
        {
            set
            {
                isKaraCell = value;
            }
            get
            {
                return isKaraCell;
            }
        }
        public bool IsContinuousLine
        {
            set
            {
                isContinuousLine = value;
            }
            get
            {
                return isContinuousLine;
            }
        }
        public SheetBorder BorderState
        {
            set
            {
                borderState = value;
            }
            get
            {
                return borderState;
            }
        }

        // カスタム ペイント実装
        protected override void Paint(Graphics graphics,
            Rectangle clipBounds, Rectangle cellBounds, int rowIndex,
            DataGridViewElementStates elementState, object value,
            object formattedValue, string errorText,
            DataGridViewCellStyle cellStyle,
            DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {
            // 常にカスタム描画を実行する。
            // 値の null 有無に描画結果（背景色/継続線/基準線）が依存すると、
            // 同期タイミング次第で装飾描画が失われるため。
            bool isSelected = (elementState & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;
            bool drawContentForeground =
                (paintParts & DataGridViewPaintParts.ContentForeground) == DataGridViewPaintParts.ContentForeground;

            // 背景の描画
            if ((paintParts & DataGridViewPaintParts.Background) ==
                DataGridViewPaintParts.Background)
            {
                using (SolidBrush cellBackground =
                    isSelected
                    ? new SolidBrush(cellStyle.SelectionBackColor)
                    : new SolidBrush(cellStyle.BackColor))
                {
                    graphics.FillRectangle(cellBackground, cellBounds);
                }
            }

            // 境界線の描画
            if ((paintParts & DataGridViewPaintParts.Border) ==
                DataGridViewPaintParts.Border)
            {
                PaintBorder(graphics, clipBounds, cellBounds, cellStyle,
                    advancedBorderStyle);
            }

            if (!drawContentForeground)
            {
                return;
            }

            // カラセル(×印)記号の描画
            if (isKaraCell)
            {
                using (Pen linepen = new Pen(Color.LightGray, 1))
                {
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y,
                        cellBounds.X + cellBounds.Width - 2,
                        cellBounds.Y + cellBounds.Height - 2);
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X + cellBounds.Width - 2,
                        cellBounds.Y,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 2);
                }
            }

            // 継続記号の描画
            if (isContinuousLine)
            {
                using (Pen linepen = new Pen(Color.LightGray, 1))
                {
                    int w = cellBounds.Right - cellBounds.Left;
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X + w / 2 - 2,
                        cellBounds.Y,
                        cellBounds.X + w / 2 - 2,
                        cellBounds.Y + cellBounds.Height - 2);
                }
            }

            //１秒毎の基準線を描画
            if ((borderState & SheetBorder.EverySec) != SheetBorder.None)
            {
                using (Pen linepen = new Pen(Color.Black, 3))
                {
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 2,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 2);
                }
            }

            //ｎコマ毎の基準線を描画
            if ((borderState & SheetBorder.EveryNFrames) != SheetBorder.None)
            {
                using (Pen linepen = new Pen(Color.LightGray, 2))
                {
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 1,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 1);
                }
            }

            //シート毎の基準線を描画
            if ((borderState & SheetBorder.EverySheet) != SheetBorder.None)
            {
                using (Pen linepen = new Pen(Color.Red, 3))
                {
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 2,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 2);
                }
            }

            // 内部領域の計算
            Rectangle baseArea = cellBounds;
            baseArea.Inflate(-2, -2);

            // 文字の描画
            // ※カラセルの場合は描画しない
            if (formattedValue is String && !isKaraCell)
            {
                TextRenderer.DrawText(graphics,
                    (string)formattedValue,
                    this.DataGridView.Font,
                    baseArea, cellStyle.ForeColor);
            }
        }

    }
}
