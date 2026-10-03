using System;
using System.Windows.Forms;

namespace AEIOU
{
    class GridCellRenderer
    {
        private readonly DataGridView _view;
        private readonly Settings _setting;
        private readonly Func<int, int, string> _cellValueGetter;

        public GridCellRenderer(DataGridView view, Settings setting, Func<int, int, string> cellValueGetter)
        {
            _view = view;
            _setting = setting;
            if (cellValueGetter == null)
            {
                throw new ArgumentNullException("cellValueGetter");
            }

            _cellValueGetter = cellValueGetter;
        }

        public void ApplyTimingCellState(DataGridViewCellPaintingEventArgs e, SheetBorder borderState, bool isContinuousLine)
        {
            TimingCell cell = _view[e.ColumnIndex, e.RowIndex] as TimingCell;
            if (cell == null)
            {
                return;
            }

            string value = _cellValueGetter(e.ColumnIndex, e.RowIndex);
            cell.BorderState = borderState;
            cell.IsKaraCell = (value == _setting.KaraCell);
            cell.IsContinuousLine = (value == "" && isContinuousLine);
        }

        public void PaintCell(DataGridViewCellPaintingEventArgs e)
        {
            e.Paint(e.ClipBounds, e.PaintParts);
            e.Handled = true;
        }

        public bool TryPaintRowHeader(DataGridViewCellPaintingEventArgs e, Action<DataGridViewCellPaintingEventArgs> rowHeaderPainter)
        {
            if (e.ColumnIndex != -1)
            {
                return false;
            }

            if (e.RowIndex < 0)
            {
                return true;
            }

            rowHeaderPainter(e);
            return true;
        }

        public void ApplyBackColor(DataGridViewCellPaintingEventArgs e, System.Drawing.Color backColor)
        {
            e.CellStyle.BackColor = backColor;
            e.CellStyle.SelectionBackColor = backColor;
        }

        public void Invalidate()
        {
            if (_view != null)
            {
                _view.Invalidate();
            }
        }
    }
}
