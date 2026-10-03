using System;
using System.Windows.Forms;

namespace AEIOU
{
    class GridSelectionService
    {
        private readonly DataGridView _view;
        private readonly Settings _setting;

        public GridSelectionService(DataGridView view, Settings setting)
        {
            _view = view;
            _setting = setting;
        }

        public Rect GetSelectedRect()
        {
            Rect r = new Rect(_setting.ColLength, _setting.RowLength, 0, 0);
            int count = _view.SelectedCells.Count;

            for (int i = 0; i < count; i++)
            {
                int rowIndex = _view.SelectedCells[i].RowIndex;
                int colIndex = _view.SelectedCells[i].ColumnIndex;
                if (rowIndex < r.Y)
                {
                    r.Y = rowIndex;
                }
                if (colIndex < r.X)
                {
                    r.X = colIndex;
                }
                if (rowIndex > r.Height)
                {
                    r.Height = rowIndex;
                }
                if (colIndex > r.Width)
                {
                    r.Width = colIndex;
                }
            }

            r.Height = r.Height - r.Y + 1;
            r.Width = r.Width - r.X + 1;

            return r;
        }

        public Rect MoveSelectionDown(Rect rect, int moveLength)
        {
            if (moveLength <= 0)
            {
                return rect;
            }

            int top = rect.Y + moveLength;
            int btm = _setting.RowLength - rect.Height;
            top = (top > btm) ? btm : top;

            return MoveSelection(rect, rect.X, top);
        }

        public Rect MoveSelection(Rect rect, int x, int y)
        {
            _view.CurrentCell = _view[x, y];

            for (int i = 0; i < rect.Height; i++)
            {
                for (int j = 0; j < rect.Width; j++)
                {
                    _view[x + j, y + i].Selected = true;
                }
            }

            rect.X = x;
            rect.Y = y;
            return rect;
        }

        public Rect MoveDown(Rect rect)
        {
            return MoveSelectionDown(rect, 1);
        }

        public Rect MoveUp(Rect rect)
        {
            int top = rect.Y - 1;
            if (top < 0)
            {
                top = 0;
            }

            return MoveSelection(rect, rect.X, top);
        }

        public Rect MoveLeft(Rect rect)
        {
            int left = rect.X - 1;
            if (left < 0)
            {
                left = 0;
            }

            return MoveSelection(rect, left, rect.Y);
        }

        public Rect MoveRight(Rect rect)
        {
            int left = rect.X + 1;
            int maxLeft = _setting.ColLength - rect.Width;
            if (left > maxLeft)
            {
                left = maxLeft;
            }

            return MoveSelection(rect, left, rect.Y);
        }
    }
}
