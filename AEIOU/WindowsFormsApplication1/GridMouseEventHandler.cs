using System;
using System.Drawing;
using System.Windows.Forms;

namespace AEIOU
{
    public interface IGridMouseService
    {
        void HandleMouseDown(DataGridViewCellMouseEventArgs e, Rect selectRange, ref Point mouseDownPoint, ref bool isRectDrag, ref bool isCtrlDragCopy);
        void HandleMouseMove(DataGridViewCellMouseEventArgs e, bool isRectDrag, DataGridView dataGridView, ref bool isCtrlDragCopy);
        void HandleMouseUp(DataGridViewCellMouseEventArgs e, DataGridView dataGridView, ref Rect selectRange, ref Point mouseDownPoint, ref bool isRectDrag, ref bool isCtrlDragCopy, ref bool isFirstEdit);
    }

    public class GridMouseEventHandler : IGridMouseService
    {
        private readonly GridViewManager gridViewManager;
        private readonly Action<Rect> copyToBuf;
        private readonly Action<Rect, bool> cutToBuf;
        private readonly Action<int, int, bool> copyToCell;
        private readonly Func<Rect> getSelectedRect;

        public GridMouseEventHandler(
            GridViewManager gridViewManager,
            Action<Rect> copyToBuf,
            Action<Rect, bool> cutToBuf,
            Action<int, int, bool> copyToCell,
            Func<Rect> getSelectedRect)
        {
            this.gridViewManager = gridViewManager;
            this.copyToBuf = copyToBuf;
            this.cutToBuf = cutToBuf;
            this.copyToCell = copyToCell;
            this.getSelectedRect = getSelectedRect;
        }

        public void HandleMouseDown(DataGridViewCellMouseEventArgs e, Rect selectRange, ref Point mouseDownPoint, ref bool isRectDrag, ref bool isCtrlDragCopy)
        {
            isCtrlDragCopy = false;

            // 左クリックかチェック
            if ((e.Button & MouseButtons.Left) != 0)
            {
                // 選択範囲内かチェック
                Rect rect = selectRange;
                int col = e.ColumnIndex;
                int row = e.RowIndex;
                if (rect.Top > row || rect.Bottom < row || rect.Left > col || rect.Right < col)
                {
                    // 初期状態に戻す
                    mouseDownPoint = new Point(-1, -1);
                    isRectDrag = false;
                }
                else if (rect.Top <= row && rect.Bottom >= row && rect.Left <= col && rect.Right >= col)
                {
                    // カレントの位置を取得
                    mouseDownPoint = new Point(col, row);
                    isRectDrag = true;
                    isCtrlDragCopy = (Control.ModifierKeys & Keys.Control) != 0;
                }
            }
        }

        public void HandleMouseMove(DataGridViewCellMouseEventArgs e, bool isRectDrag, DataGridView dataGridView, ref bool isCtrlDragCopy)
        {
            if (isRectDrag)
            {
                isCtrlDragCopy = (Control.ModifierKeys & Keys.Control) != 0;

                // 描画更新(範囲描画のため)
                dataGridView.Invalidate();
            }
        }

        public void HandleMouseUp(DataGridViewCellMouseEventArgs e, DataGridView dataGridView, ref Rect selectRange, ref Point mouseDownPoint, ref bool isRectDrag, ref bool isCtrlDragCopy, ref bool isFirstEdit)
        {
            // ドラッグ中だった場合は 選択範囲を修正し移動（又はコピー）処理を行う
            if (isRectDrag && dataGridView.CurrentCell != null)
            {
                //選択範囲を解除
                dataGridView.ClearSelection();

                //カレントセルは選択しておく
                dataGridView.CurrentCell.Selected = true;

                bool executeCopy = isCtrlDragCopy || ((Control.ModifierKeys & Keys.Control) != 0);
                if (executeCopy)
                {
                    //選択元をコピー＆ペースト
                    int col = dataGridView.CurrentCell.ColumnIndex - (mouseDownPoint.X - selectRange.X);
                    int row = dataGridView.CurrentCell.RowIndex - (mouseDownPoint.Y - selectRange.Y);
                    gridViewManager.BeginGroup("選択元をコピー＆ペースト");
                    copyToBuf(selectRange);
                    copyToCell(col, row, false);
                    gridViewManager.EndGroup();
                }
                else
                {
                    //選択元をカット＆ペースト
                    int col = dataGridView.CurrentCell.ColumnIndex - (mouseDownPoint.X - selectRange.X);
                    int row = dataGridView.CurrentCell.RowIndex - (mouseDownPoint.Y - selectRange.Y);
                    gridViewManager.BeginGroup("選択元をカット＆ペースト");
                    cutToBuf(selectRange, false);
                    copyToCell(col, row, false);
                    gridViewManager.EndGroup();
                }
            }
            else
            {
                // Ctrlキー同時押しの抑止
                if ((Control.ModifierKeys & Keys.Control) != 0)
                {
                    // 選択領域を解除 (カレントフレームのみ選択する)
                    dataGridView.ClearSelection();
                    if(dataGridView.CurrentCell != null)
                    {
                         dataGridView.CurrentCell.Selected = true;
                    }
                }
            }

            // 初期状態に戻す
            mouseDownPoint = new Point(-1, -1);
            isRectDrag = false;
            isCtrlDragCopy = false;
            isFirstEdit = true;

            // 選択範囲を保存
            selectRange = getSelectedRect();

            // 描画更新(範囲描画のため)
            dataGridView.Invalidate();
        }
    }
}
