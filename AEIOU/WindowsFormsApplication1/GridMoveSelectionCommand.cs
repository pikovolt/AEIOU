using System;
using System.Windows.Forms;

namespace AEIOU
{
    class GridMoveSelectionCommand
    {
        private readonly DataGridView view;
        private readonly Settings setting;
        private readonly GridSelectionService gridSelectionService;
        private readonly GridScrollService gridScrollService;

        public GridMoveSelectionCommand(
            DataGridView view,
            Settings setting,
            GridSelectionService gridSelectionService,
            GridScrollService gridScrollService)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            if (setting == null)
            {
                throw new ArgumentNullException("setting");
            }

            if (gridSelectionService == null)
            {
                throw new ArgumentNullException("gridSelectionService");
            }

            if (gridScrollService == null)
            {
                throw new ArgumentNullException("gridScrollService");
            }

            this.view = view;
            this.setting = setting;
            this.gridSelectionService = gridSelectionService;
            this.gridScrollService = gridScrollService;
        }

        public Rect HandleEnter(Rect selectRange, Func<int> moveLengthProvider)
        {
            if (moveLengthProvider == null)
            {
                throw new ArgumentNullException("moveLengthProvider");
            }

            Rect rect = gridSelectionService.GetSelectedRect();
            view.ClearSelection();

            int len = moveLengthProvider();
            if (len > 0)
            {
                selectRange = gridSelectionService.MoveSelectionDown(rect, len);
            }

            gridScrollService.ScrollForward(selectRange);
            return selectRange;
        }

        public void HandlePageUp()
        {
            gridScrollService.ScrollVertical(-view.DisplayedRowCount(true));
        }

        public void HandlePageDown()
        {
            gridScrollService.ScrollVertical(view.DisplayedRowCount(true));
        }

        public Rect HandleHome(Rect selectRange)
        {
            view.ClearSelection();
            return gridSelectionService.MoveSelection(selectRange, selectRange.X, 0);
        }

        public Rect HandleLeftArrow(Rect selectRange, bool isShiftPressed)
        {
            if (isShiftPressed)
            {
                Rect rect = gridSelectionService.GetSelectedRect();
                if (rect.Width > 1)
                {
                    for (int i = 0; i < rect.Height; i++)
                    {
                        view[rect.Right, rect.Y + i].Selected = false;
                    }

                    rect.Width--;
                    return rect;
                }

                return selectRange;
            }

            return gridSelectionService.MoveLeft(selectRange);
        }

        public Rect HandleUpArrow(Rect selectRange, bool isShiftPressed)
        {
            if (isShiftPressed)
            {
                return HandleDivide(selectRange);
            }

            return gridSelectionService.MoveUp(selectRange);
        }

        public Rect HandleRightArrow(Rect selectRange, bool isShiftPressed)
        {
            if (isShiftPressed)
            {
                Rect rect = gridSelectionService.GetSelectedRect();
                if (rect.Right + 1 < setting.ColLength)
                {
                    for (int i = 0; i < rect.Height; i++)
                    {
                        view[rect.Right + 1, rect.Y + i].Selected = true;
                    }

                    rect.Width++;
                    return rect;
                }

                return selectRange;
            }

            return gridSelectionService.MoveRight(selectRange);
        }

        public Rect HandleDownArrow(Rect selectRange, bool isShiftPressed)
        {
            if (isShiftPressed)
            {
                return HandleMultiply(selectRange);
            }

            return gridSelectionService.MoveDown(selectRange);
        }

        public Rect HandleMultiply(Rect selectRange)
        {
            Rect rect = gridSelectionService.GetSelectedRect();
            if (rect.Bottom + 1 < setting.RowLength)
            {
                for (int j = 0; j < rect.Width; j++)
                {
                    view[rect.X + j, rect.Bottom + 1].Selected = true;
                }

                rect.Height++;
                return rect;
            }

            return selectRange;
        }

        public Rect HandleDivide(Rect selectRange)
        {
            Rect rect = gridSelectionService.GetSelectedRect();
            if (rect.Height > 1)
            {
                for (int j = 0; j < rect.Width; j++)
                {
                    view[rect.X + j, rect.Bottom].Selected = false;
                }

                rect.Height--;
                return rect;
            }

            return selectRange;
        }
    }
}