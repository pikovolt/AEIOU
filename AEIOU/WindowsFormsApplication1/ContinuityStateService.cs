namespace AEIOU
{
    class ContinuityStateService
    {
        private readonly Settings _setting;
        private readonly System.Func<int, int, string> _getCellValue;
        private bool[,] _continuityFlags;

        public ContinuityStateService(Settings setting, System.Func<int, int, string> getCellValue)
        {
            _setting = setting;
            _getCellValue = getCellValue;
            _continuityFlags = new bool[0, 0];
        }

        public void Reinitialize(int columnCount, int rowCount)
        {
            if (columnCount < 0) columnCount = 0;
            if (rowCount < 0) rowCount = 0;

            _continuityFlags = new bool[columnCount, rowCount];
            for (int col = 0; col < columnCount; col++)
            {
                RecalculateColumn(col, rowCount);
            }
        }

        public void RecalculateColumn(int col, int rowCount)
        {
            if (_continuityFlags == null) return;
            if (col < 0 || col >= _continuityFlags.GetLength(0)) return;

            bool hasTiming = false;
            int maxRow = _continuityFlags.GetLength(1);
            int limit = (rowCount < maxRow) ? rowCount : maxRow;

            for (int row = 0; row < limit; row++)
            {
                string value = _getCellValue(col, row) ?? string.Empty;
                if (value == string.Empty)
                {
                    // 直前状態を維持
                }
                else if (value == _setting.KaraCell)
                {
                    hasTiming = false;
                }
                else
                {
                    hasTiming = true;
                }

                _continuityFlags[col, row] = hasTiming;
            }

            for (int row = limit; row < maxRow; row++)
            {
                _continuityFlags[col, row] = false;
            }
        }

        public bool GetContinuityFlag(int col, int row)
        {
            if (_continuityFlags == null) return false;
            if (col < 0 || row < 0) return false;
            if (col >= _continuityFlags.GetLength(0) || row >= _continuityFlags.GetLength(1)) return false;

            return _continuityFlags[col, row];
        }
    }
}
