using System;
using System.Drawing;
using System.Windows.Forms;

namespace AEIOU
{
    class GridScrollService
    {
        private readonly DataGridView _view;
        private readonly Settings _setting;

        public GridScrollService(DataGridView view, Settings setting)
        {
            _view = view;
            _setting = setting;
        }

        public void ScrollForward(Rect selectRange)
        {
            Rectangle client = _view.ClientRectangle;
            int height = _view.Rows[0].Height;
            int rowTop = _view.FirstDisplayedScrollingRowIndex;
            int currentRow = _view.CurrentCell.RowIndex - rowTop;
            int rowCount = (client.Height - (height - 1)) / height;
            int border = (int)(rowCount * (2.0 / 3.0));
            if (border < currentRow)
            {
                int forwardCount = currentRow - border;
                int top = selectRange.Top + forwardCount;
                int btm = _setting.RowLength - selectRange.Height;
                if (top > btm)
                {
                    _view.FirstDisplayedScrollingRowIndex = btm;
                }
                else
                {
                    _view.FirstDisplayedScrollingRowIndex += forwardCount;
                }
            }
        }

        public void ScrollRowBackward(int keyValue)
        {
            int moveSize;
            if (_setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.CTRLKey))
            {
                moveSize = _setting.Fps * _setting.SheetSec;
            }
            else
            {
                moveSize = _setting.Fps;
            }

            int row = _view.FirstDisplayedScrollingRowIndex;
            int newRow = row - moveSize;
            if ((newRow % moveSize) != 0)
            {
                newRow += moveSize - (newRow % moveSize);
            }

            if (newRow >= 0)
            {
                _view.FirstDisplayedScrollingRowIndex = newRow;
                _view.CurrentCell = _view[_view.CurrentCell.ColumnIndex, newRow];
            }
        }

        public void ScrollRowForward(int keyValue)
        {
            int moveSize;
            if (_setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.CTRLKey))
            {
                moveSize = _setting.Fps * _setting.SheetSec;
            }
            else
            {
                moveSize = _setting.Fps;
            }

            int row = _view.FirstDisplayedScrollingRowIndex;
            int newRow = row + moveSize;
            newRow -= newRow % moveSize;

            if (newRow < _setting.RowLength)
            {
                _view.FirstDisplayedScrollingRowIndex = newRow;
                _view.CurrentCell = _view[_view.CurrentCell.ColumnIndex, newRow];
            }
        }

        public void ScrollVertical(int offsetRows)
        {
            if (offsetRows == 0)
            {
                return;
            }

            int currentTop = _view.FirstDisplayedScrollingRowIndex;
            int maxTop = _setting.RowLength - 1;
            int newTop = currentTop + offsetRows;

            if (newTop < 0)
            {
                newTop = 0;
            }
            if (newTop > maxTop)
            {
                newTop = maxTop;
            }

            _view.FirstDisplayedScrollingRowIndex = newTop;

            int currentCol = _view.CurrentCell.ColumnIndex;
            int currentRow = _view.CurrentCell.RowIndex + offsetRows;
            if (currentRow < 0)
            {
                currentRow = 0;
            }
            if (currentRow >= _setting.RowLength)
            {
                currentRow = _setting.RowLength - 1;
            }

            _view.CurrentCell = _view[currentCol, currentRow];
        }
    }
}
