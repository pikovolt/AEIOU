using System;
using System.Windows.Forms;

namespace AEIOU
{
    class GridValueInputCommand
    {
        private readonly DataGridView view;
        private readonly Settings setting;
        private readonly IGridCellValueService gridCellValueService;
        private readonly GridSelectionService gridSelectionService;
        private readonly GridScrollService gridScrollService;

        public GridValueInputCommand(
            DataGridView view,
            Settings setting,
            IGridCellValueService gridCellValueService,
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

            if (gridCellValueService == null)
            {
                throw new ArgumentNullException("gridCellValueService");
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
            this.gridCellValueService = gridCellValueService;
            this.gridSelectionService = gridSelectionService;
            this.gridScrollService = gridScrollService;
        }

        public bool HandleNumber(Rect selectRange, int keyValue, bool isFirstEdit)
        {
            int normalizedKey = keyValue & 0x0ff;
            if (normalizedKey >= 96 && normalizedKey <= 105)
            {
                normalizedKey -= 96;
            }
            else if (normalizedKey >= 48 && normalizedKey <= 57)
            {
                normalizedKey -= 48;
            }
            else
            {
                return false;
            }

            gridCellValueService.InsertNumber(selectRange, normalizedKey, isFirstEdit, setting.IsAlwaysAppend);
            return true;
        }

        public Rect HandleAdd(Rect selectRange, Func<int> moveLengthProvider)
        {
            gridCellValueService.IncrementValue(selectRange, setting.KaraCell);
            return MoveAfterInput(selectRange, moveLengthProvider);
        }

        public Rect HandleSubtract(Rect selectRange, Func<int> moveLengthProvider)
        {
            gridCellValueService.DecrementValue(selectRange, setting.KaraCell);
            return MoveAfterInput(selectRange, moveLengthProvider);
        }

        public Rect HandleDecimal(Rect selectRange, Func<int> moveLengthProvider)
        {
            gridCellValueService.InsertEmptyCell(selectRange, setting.KaraCell);
            if (setting.IsKaraNoMove)
            {
                return selectRange;
            }

            return MoveAfterInput(selectRange, moveLengthProvider);
        }

        private Rect MoveAfterInput(Rect selectRange, Func<int> moveLengthProvider)
        {
            if (moveLengthProvider == null)
            {
                throw new ArgumentNullException("moveLengthProvider");
            }

            view.ClearSelection();
            int len = moveLengthProvider();
            if (len > 0)
            {
                selectRange = gridSelectionService.MoveSelectionDown(selectRange, len);
            }

            gridScrollService.ScrollForward(selectRange);
            return selectRange;
        }
    }
}
