using System;
using System.Windows.Forms;

namespace AEIOU
{
    public interface IGridShortcutHandler
    {
        void OnBackSpace();
        void OnEnter();
        void OnPageUp(int keyValue);
        void OnPageDown(int keyValue);
        void OnHome();
        void OnLeftArrow(bool isShiftPressed);
        void OnUpArrow(bool isShiftPressed);
        void OnRightArrow(bool isShiftPressed);
        void OnDownArrow(bool isShiftPressed);
        void OnInsert();
        void OnDelete(int keyValue);
        void OnNumberKey(int keyValue, int keyCode);
        void OnJOrKKey(int keyValue);
        void OnMultiplyKey();
        void OnAddKey();
        void OnSubtractKey();
        void OnDivideKey();
        void OnDecimalKey();
    }

    public class GridShortcutRouter
    {
        private readonly IGridShortcutHandler handler;

        public GridShortcutRouter(IGridShortcutHandler handler)
        {
            this.handler = handler;
        }

        public bool Route(int keyValue, int keyCode, bool isShiftPressed)
        {
            switch (keyValue & 0x0ff)
            {
                case 8:     // BackSpace
                    handler.OnBackSpace();
                    return true;

                case 13:    // Enter
                    handler.OnEnter();
                    return true;

                case 33:    // PageUp
                    handler.OnPageUp(keyValue);
                    return true;

                case 34:    // PageDown
                    handler.OnPageDown(keyValue);
                    return true;

                case 36:    // Home
                    handler.OnHome();
                    return true;

                case 37:    // ←
                    handler.OnLeftArrow(isShiftPressed);
                    return true;

                case 38:    //↑
                    handler.OnUpArrow(isShiftPressed);
                    return true;

                case 39:    // →
                    handler.OnRightArrow(isShiftPressed);
                    return true;

                case 40:    // ↓
                    handler.OnDownArrow(isShiftPressed);
                    return true;

                case 45:    // Insert
                    handler.OnInsert();
                    return true;

                case 46:    // Delete
                    handler.OnDelete(keyValue);
                    return true;

                case 48:    // 0-9(Full-key)
                case 49:
                case 50:
                case 51:
                case 52:
                case 53:
                case 54:
                case 55:
                case 56:
                case 57:
                case 96:    // 0-9(10key)
                case 97:
                case 98:
                case 99:
                case 100:
                case 101:
                case 102:
                case 103:
                case 104:
                case 105:
                    handler.OnNumberKey(keyValue, keyCode);
                    return true;

                case 74:    // J
                case 75:    // K
                    handler.OnJOrKKey(keyValue);
                    return true;

                case 106:   // '*'(10key)
                    handler.OnMultiplyKey();
                    return true;

                case 107:   // '+'(10key)
                case 187:   // '+'(full-key)
                    handler.OnAddKey();
                    return true;

                case 109:   // '-'(10key)
                case 189:   // '-'(full-key)
                    handler.OnSubtractKey();
                    return true;

                case 111:   // '/'(10key)
                    handler.OnDivideKey();
                    return true;

                case 110:   // '.'(10key)
                case 190:   // '.'(full-key)
                    handler.OnDecimalKey();
                    return true;

                default:
                    return false;
            }
        }
    }
}
