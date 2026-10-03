using System;
using System.Windows.Forms;

namespace AEIOU
{
    public partial class Form1 : Form, IGridShortcutHandler
    {
        public void OnBackSpace()
        {
            if (isFirstEdit || setting.IsAlwaysAppend)
            {
                // Delete selection
                isCellEdit = deleteRect_with_backspace(isCellEdit);
            }
            else
            {
                // 1 char delete logic
                for (int row = selectRange.Top; row <= selectRange.Bottom; row++)
                {
                    for (int col = selectRange.Left; col <= selectRange.Right; col++)
                    {
                        string str = GetCellValue(col, row);
                        if (!string.IsNullOrEmpty(str) && str.Length > 0 && str != setting.KaraCell)
                        {
                            str = str.Substring(0, str.Length - 1);
                            gridCellValueService.ApplyValue(new Rect(col, row, 1, 1), str);
                            isCellEdit = true;
                        }
                    }
                }
            }
        }

        public void OnEnter()
        {
            selectRange = gridMoveSelectionCommand.HandleEnter(selectRange, cursorMoveWithNakaNuki);
            isFirstEdit = true;
        }
        public void OnPageUp(int keyValue) { gridMoveSelectionCommand.HandlePageUp(); }
        public void OnPageDown(int keyValue) { gridMoveSelectionCommand.HandlePageDown(); }
        public void OnHome() { selectRange = gridMoveSelectionCommand.HandleHome(selectRange); }
        public void OnLeftArrow(bool isShiftPressed)
        {
            selectRange = gridMoveSelectionCommand.HandleLeftArrow(selectRange, isShiftPressed);
        }

        public void OnUpArrow(bool isShiftPressed)
        {
            selectRange = gridMoveSelectionCommand.HandleUpArrow(selectRange, isShiftPressed);
        }

        public void OnRightArrow(bool isShiftPressed)
        {
            selectRange = gridMoveSelectionCommand.HandleRightArrow(selectRange, isShiftPressed);
        }

        public void OnDownArrow(bool isShiftPressed)
        {
            selectRange = gridMoveSelectionCommand.HandleDownArrow(selectRange, isShiftPressed);
        }
        
        public void OnInsert() 
        { 
            insertToAllCell(selectRange.Top, selectRange.Height);
            calcNakanukiRange(true, selectRange.Top, selectRange.Height);
            calcKiribariRange(true, selectRange.Top, selectRange.Height);
            flushUndoHistory();
            selectRange = gridSelectionService.MoveSelectionDown(selectRange, selectRange.Height);
        }

        public void OnDelete(int keyValue) 
        { 
            if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.SHIFTKey))
            {
                // Shift+Delete: 範囲削除
                cutToAllCell(selectRange.Top, selectRange.Height);
                calcNakanukiRange(false, selectRange.Top, selectRange.Height);
                calcKiribariRange(false, selectRange.Top, selectRange.Height);
                flushUndoHistory();
                int top = selectRange.Y - selectRange.Height;
                if (top < 0)
                {
                    top = 0;
                }
                selectRange = gridSelectionService.MoveSelection(selectRange, selectRange.X, top);
            }
            else if (setting.keys.checkShiftBeforeConvertion(keyValue, CombinationKeyState.None))
            {
                // Delete: 選択範囲の内容削除（移動なし）
                deleteRect(selectRange);
            }
        }

        public void OnNumberKey(int keyValue, int keyCode)
        {
            if (gridValueInputCommand.HandleNumber(selectRange, keyValue, isFirstEdit))
            {
                isCellEdit = true;
            }
        }

        public void OnJOrKKey(int keyCode)
        {
            int value = selectRange.Top;
            int col, row;
            if (keyCode == 74) // J
            {
                for (row = selectRange.Top - 1; row >= 0 && value == selectRange.Top; row--)
                {
                    for (col = selectRange.Left; col <= selectRange.Right; col++)
                    {
                        if (GetCellValue(col, row) == "") continue;
                        value = row;
                        break;
                    }
                }
            }
            if (keyCode == 75) // K
            {
                for (row = selectRange.Bottom + 1; row < setting.RowLength && value == selectRange.Top; row++)
                {
                    for (col = selectRange.Left; col <= selectRange.Right; col++)
                    {
                        if (GetCellValue(col, row) == "") continue;
                        value = row;
                        break;
                    }
                }
            }
            if (value != selectRange.Top)
            {
                this.dataGridView1.ClearSelection();
                selectRange = gridSelectionService.MoveSelection(selectRange, selectRange.X, value);
            }
        }

        public void OnMultiplyKey() 
        { 
            selectRange = gridMoveSelectionCommand.HandleMultiply(selectRange);
        }
        public void OnAddKey()
        {
            selectRange = gridValueInputCommand.HandleAdd(selectRange, cursorMoveWithNakaNuki);
        }

        public void OnSubtractKey()
        {
            selectRange = gridValueInputCommand.HandleSubtract(selectRange, cursorMoveWithNakaNuki);
        }
        public void OnDivideKey() 
        { 
            selectRange = gridMoveSelectionCommand.HandleDivide(selectRange);
        }
        
        public void OnDecimalKey() 
        {
            selectRange = gridValueInputCommand.HandleDecimal(selectRange, cursorMoveWithNakaNuki);
            isCellEdit = true;
        }
    }
}
