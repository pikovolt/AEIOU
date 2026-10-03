using System;

namespace AEIOU
{
    public interface IGridCellValueService
    {
        void InsertNumber(Rect selectRange, int key, bool isFirstEdit, bool alwaysAppend);
        void IncrementValue(Rect selectRange, string karaCellSymbol);
        void DecrementValue(Rect selectRange, string karaCellSymbol);
        void InsertEmptyCell(Rect selectRange, string karaCellSymbol);
    }

    public class GridCellValueService : IGridCellValueService
    {
        private readonly GridViewManager gridViewManager;
        private readonly Func<int, int, string> getCellValue;
        private readonly Func<int, int, bool> checkCellValue;
        private readonly Action<int> incrementCellUsedCount;

        public GridCellValueService(
            GridViewManager gridViewManager,
            Func<int, int, string> getCellValue,
            Func<int, int, bool> checkCellValue,
            Action<int> incrementCellUsedCount)
        {
            this.gridViewManager = gridViewManager;
            this.getCellValue = getCellValue;
            this.checkCellValue = checkCellValue;
            this.incrementCellUsedCount = incrementCellUsedCount;
        }

        public void InsertNumber(Rect selectRange, int key, bool isFirstEdit, bool alwaysAppend)
        {
            if (key < 0 || key > 9)
            {
                return;
            }

            string digitText = ((char)('0' + key)).ToString();

            gridViewManager.BeginGroup("入力");

            for (int i = 0; i < selectRange.Width; i++)
            {
                string newValue = getCellValue(selectRange.X + i, selectRange.Y);

                // 入力前に情報が入っているか確認
                if (!checkCellValue(selectRange.X + i, selectRange.Y))
                {
                    // 空白の場合は使用状況を修正
                    incrementCellUsedCount(selectRange.X + i);
                }

                // 初回フラグが立ち, 尚且つ "常に追加"が未チェックの場合のみ 初回編集を上書きにする
                if (isFirstEdit && !alwaysAppend)
                {
                    // セルに値を設定(初回編集)
                    newValue = digitText;
                }
                else
                {
                    // セルに値を設定(継続編集)
                    newValue += digitText;
                }

                // アンドゥ情報の記録
                var operation = new SetValueOperation(selectRange.Y, selectRange.X + i, newValue);
                gridViewManager.ExecuteOperation(operation);
            }

            gridViewManager.EndGroup();
        }

        public void IncrementValue(Rect selectRange, string karaCellSymbol)
        {
             // 入力前に情報が入っているか確認
            if (!checkCellValue(selectRange.X, selectRange.Y))
            {
                // 空白の場合は使用状況を修正
                incrementCellUsedCount(selectRange.X);
            }

            // 手前の入力を検索
            int key = 1;
            for (int i = selectRange.Y - 1; i >= 0; i--)
            {
                // 空白と空セルは無視
                string str = getCellValue(selectRange.X, i);
                if (string.IsNullOrEmpty(str)) continue;
                if (str == karaCellSymbol) return; // カラセルが入力されていたら、処理を中断
                
                if (int.TryParse(str, out int val))
                {
                    key = val + 1;
                    break;
                }
            }
            
            var operation = new SetValueOperation(selectRange.Y, selectRange.X, key.ToString());
            gridViewManager.ExecuteOperation(operation);
        }
        
        public void DecrementValue(Rect selectRange, string karaCellSymbol)
        {
             // 入力前に情報が入っているか確認
            if (!checkCellValue(selectRange.X, selectRange.Y))
            {
                // 空白の場合は使用状況を修正
                incrementCellUsedCount(selectRange.X);
            }

            // 手前の入力を検索
            int key = 1;
            for (int i = selectRange.Y - 1; i >= 0; i--)
            {
                // 空白と空セルは無視
                string str = getCellValue(selectRange.X, i);
                if (string.IsNullOrEmpty(str)) continue;
                if (str == karaCellSymbol) return; // カラセルが入力されていたら、処理を中断
                
                if (int.TryParse(str, out int val))
                {
                    key = val - 1;
                    if (key < 1) key = 1; // 1未満にはならない前提
                    break;
                }
            }
            
            var operation = new SetValueOperation(selectRange.Y, selectRange.X, key.ToString());
            gridViewManager.ExecuteOperation(operation);
        }
        
        public void InsertEmptyCell(Rect selectRange, string karaCellSymbol)
        {
             gridViewManager.BeginGroup("空セルの入力");

            // 値の入力
            for (int i = selectRange.Left; i <= selectRange.Right; i++)
            {
                // 入力前に情報が入っているか確認
                if (!checkCellValue(i, selectRange.Y))
                {
                    // 空白の場合は使用状況を修正
                    incrementCellUsedCount(i);
                }

                // アンドゥ情報の記録
                var operation = new SetValueOperation(selectRange.Y, i, karaCellSymbol);
                gridViewManager.ExecuteOperation(operation);
            }

            gridViewManager.EndGroup();
        }
        
        public void ApplyValue(Rect selectRange, string valueToApply)
        {
            // 入力前に情報が入っているか確認
            if (!checkCellValue(selectRange.X, selectRange.Y))
            {
                // 空白の場合は使用状況を修正
                incrementCellUsedCount(selectRange.X);
            }

            // アンドゥ情報の記録
            var operation = new SetValueOperation(selectRange.Y, selectRange.X, valueToApply);
            gridViewManager.ExecuteOperation(operation);
        }
    }
}
