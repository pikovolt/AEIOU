using System;
using System.Windows.Forms;

namespace AEIOU
{
    enum GridDispatchResult
    {
        None,
        MenuShortcut,
        GridCommand
    }

    class GridKeyCommandDispatcher
    {
        private readonly GridInputInterpreter inputInterpreter;
        private readonly GridShortcutRouter shortcutRouter;
        private readonly Func<ToolStripItemCollection, Keys, bool> shortcutExecutor;

        internal GridKeyCommandDispatcher(
            GridInputInterpreter inputInterpreter,
            GridShortcutRouter shortcutRouter,
            Func<ToolStripItemCollection, Keys, bool> shortcutExecutor)
        {
            if (inputInterpreter == null)
            {
                throw new ArgumentNullException("inputInterpreter");
            }

            if (shortcutRouter == null)
            {
                throw new ArgumentNullException("shortcutRouter");
            }

            if (shortcutExecutor == null)
            {
                throw new ArgumentNullException("shortcutExecutor");
            }

            this.inputInterpreter = inputInterpreter;
            this.shortcutRouter = shortcutRouter;
            this.shortcutExecutor = shortcutExecutor;
        }

        public GridDispatchResult Dispatch(KeyEventArgs e, ToolStripItemCollection items, out int convertedKeyValue)
        {
            if (e == null)
            {
                throw new ArgumentNullException("e");
            }

            convertedKeyValue = 0;

            if (inputInterpreter.TryHandleShortcut(e, items, shortcutExecutor))
            {
                return GridDispatchResult.MenuShortcut;
            }

            convertedKeyValue = inputInterpreter.ConvertKeyValue(e);
            if (shortcutRouter.Route(convertedKeyValue, e.KeyValue, e.Shift))
            {
                return GridDispatchResult.GridCommand;
            }

            return GridDispatchResult.None;
        }
    }
}
