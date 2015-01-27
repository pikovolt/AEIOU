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
            // 値が入っている場合のみ、カスタム描画を実行
            if (value != null)
            {

                // 背景の描画
                if ((paintParts & DataGridViewPaintParts.Background) ==
                    DataGridViewPaintParts.Background)
                {
                    SolidBrush cellBackground =
                        (elementState == DataGridViewElementStates.Selected)
                        ? new SolidBrush(cellStyle.SelectionBackColor)
                        : new SolidBrush(cellStyle.BackColor);
                    graphics.FillRectangle(cellBackground, cellBounds);
                    cellBackground.Dispose();
                }

                // 境界線の描画
                if ((paintParts & DataGridViewPaintParts.Border) ==
                    DataGridViewPaintParts.Border)
                {
                    PaintBorder(graphics, clipBounds, cellBounds, cellStyle,
                        advancedBorderStyle);
                }

                // カラセル(×印)記号の描画
                if (isKaraCell)
                {
                    Pen linepen = new Pen(Color.LightGray, 1);
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

                // 継続記号の描画
                if (isContinuousLine)
                {
                    Pen linepen = new Pen(Color.LightGray, 1);
                    int w = cellBounds.Right - cellBounds.Left;
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X + w / 2 - 2,
                        cellBounds.Y,
                        cellBounds.X + w / 2 - 2,
                        cellBounds.Y + cellBounds.Height - 2);
                }

                //１秒毎の基準線を描画
                if ((borderState & SheetBorder.EverySec) != SheetBorder.None)
                {
                    Pen linepen = new Pen(Color.Black, 3);
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 2,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 2);
                }

                //ｎコマ毎の基準線を描画
                if ((borderState & SheetBorder.EveryNFrames) != SheetBorder.None)
                {
                    Pen linepen = new Pen(Color.LightGray, 2);
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 1,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 1);
                }

                //シート毎の基準線を描画
                if ((borderState & SheetBorder.EverySheet) != SheetBorder.None)
                {
                    Pen linepen = new Pen(Color.Red, 3);
                    graphics.DrawLine(
                        linepen,
                        cellBounds.X,
                        cellBounds.Y + cellBounds.Height - 2,
                        cellBounds.X + cellBounds.Width,
                        cellBounds.Y + cellBounds.Height - 2);
                }

                // 内部領域の計算
                Rectangle baseArea = cellBounds;
                baseArea.Inflate(-2, -2);

                // 文字の描画
                // ※カラセルの場合は描画しない
                if (this.FormattedValue is String && !isKaraCell)
                {
                    TextRenderer.DrawText(graphics,
                        (string)this.FormattedValue,
                        this.DataGridView.Font,
                        baseArea, cellStyle.ForeColor);
                }
            }
            else
            {
                // ベースクラスでの描画
                base.Paint(graphics, clipBounds, cellBounds, rowIndex,
                    elementState, value, formattedValue, errorText,
                    cellStyle, advancedBorderStyle, paintParts);
            }
        }

    }
}
