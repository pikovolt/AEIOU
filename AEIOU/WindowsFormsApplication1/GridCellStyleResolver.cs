using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace AEIOU
{
    class GridCellStyleResolver
    {
        public Color ResolveBackColor(
            DataGridView view,
            DataGridViewCellPaintingEventArgs e,
            ColorDefinitions gridPalette,
            int[] aryCellUsedCount,
            List<Range> delRange,
            List<Range> addRange,
            bool isRectDrag,
            Rect selectRange,
            Point mouseDownPoint)
        {
            bool bUsed = (aryCellUsedCount[e.ColumnIndex] > 0);
            bool bActive = (e.ColumnIndex == view.CurrentCell.ColumnIndex);

            bool bNuki = false;
            foreach (Range nuki in delRange)
            {
                if (nuki.Top <= e.RowIndex && nuki.Bottom >= e.RowIndex)
                {
                    bNuki = true;
                }
            }

            bool bKiribari = false;
            foreach (Range kiribari in addRange)
            {
                if (kiribari.Top <= e.RowIndex && kiribari.Bottom >= e.RowIndex)
                {
                    bKiribari = true;
                }
            }

            Color bgColor = new Color();
            if (bNuki)
            {
                bgColor = gridPalette.Nakanuki;
            }
            else if (bKiribari)
            {
                bgColor = gridPalette.Harikomi;
            }
            else
            {
                switch (view.Columns[e.ColumnIndex].DisplayIndex % 2)
                {
                    case 0:
                        if (bActive) bgColor = gridPalette.BgCell1A;
                        else bgColor = (bUsed) ? gridPalette.BgCell1R : gridPalette.BgCell1;
                        break;
                    case 1:
                        if (bActive) bgColor = gridPalette.BgCell2A;
                        else bgColor = (bUsed) ? gridPalette.BgCell2R : gridPalette.BgCell2;
                        break;
                }
            }

            bool bSelected =
                (e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;

            bool bDragSource = false;
            if (isRectDrag &&
                (selectRange.Top <= e.RowIndex && selectRange.Bottom >= e.RowIndex &&
                 selectRange.Left <= e.ColumnIndex && selectRange.Right >= e.ColumnIndex))
            {
                bDragSource = true;
            }

            bool bMovingRange = false;
            if (isRectDrag)
            {
                Point offset = new Point(mouseDownPoint.X - selectRange.X, mouseDownPoint.Y - selectRange.Y);
                int col = view.CurrentCell.ColumnIndex;
                int row = view.CurrentCell.RowIndex;
                int w = selectRange.Width - 1;
                int h = selectRange.Height - 1;
                if (!(((col - offset.X) > e.ColumnIndex) ||
                      ((col + w - offset.X) < e.ColumnIndex) ||
                      ((row - offset.Y) > e.RowIndex) ||
                      ((row + h - offset.Y) < e.RowIndex)))
                {
                    bMovingRange = true;
                }
            }

            if (isRectDrag)
            {
                if (bMovingRange)
                {
                    bgColor = gridPalette.Selected;
                }
                else if (bDragSource)
                {
                    bgColor = Color.DarkGray;
                }
            }
            else if (bSelected)
            {
                bgColor = gridPalette.Selected;
            }

            return bgColor;
        }
    }
}
