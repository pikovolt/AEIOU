using System;
using System.Windows.Forms;

namespace AEIOU
{
    class GridInputInterpreter
    {
        private readonly KeyBinds _keyBinds;

        public GridInputInterpreter(KeyBinds keyBinds)
        {
            if (keyBinds == null)
            {
                throw new ArgumentNullException("keyBinds");
            }

            _keyBinds = keyBinds;
        }

        public int ConvertKeyValue(KeyEventArgs e)
        {
            return _keyBinds.convKey(e.KeyValue, e.Alt, e.Control, e.Shift);
        }

        public bool TryHandleShortcut(KeyEventArgs e, ToolStripItemCollection items, Func<ToolStripItemCollection, Keys, bool> shortcutExecutor)
        {
            if (shortcutExecutor == null)
            {
                throw new ArgumentNullException("shortcutExecutor");
            }

            int keyValue = ConvertKeyValue(e);
            Keys convertedKeyData = (Keys)keyValue;
            if (!shortcutExecutor(items, convertedKeyData) && !shortcutExecutor(items, e.KeyData))
            {
                return false;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
            return true;
        }
    }
}
