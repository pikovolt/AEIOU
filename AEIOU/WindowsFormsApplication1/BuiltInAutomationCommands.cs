using System;
using System.Collections.Generic;
using System.Globalization;
using global::AEIOU.Automation;

namespace AEIOU
{
    public sealed class ReplaceCommand : IAutomationCommand
    {
        public const string CommandId = "builtin.replace";
        public const string BeforeParameter = "before";
        public const string AfterParameter = "after";

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "置換",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion,
                    new[]
                    {
                        Parameter(BeforeParameter, "置換前"),
                        Parameter(AfterParameter, "置換後")
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            string before;
            string after;
            if (!request.Parameters.TryGetValue(BeforeParameter, out before) || before.Length == 0)
                return AutomationResult.Failure("変換前指定がない.");
            if (!request.Parameters.TryGetValue(AfterParameter, out after) || after.Length == 0)
                return AutomationResult.Failure("変換後指定がない.");

            List<AutomationChange> changes = new List<AutomationChange>();
            int column = request.Selection.Left;
            foreach (AutomationCell cell in request.Cells)
                if (cell.Column == column && cell.Value.Length != 0 && cell.Value == before)
                    changes.Add(new AutomationChange(cell.Row, cell.Column, after));
            return AutomationResult.Success(changes, null);
        }

        private static AutomationParameterDefinition Parameter(string id, string name)
        {
            return new AutomationParameterDefinition(id, name, AutomationParameterType.String,
                String.Empty, true, null, null, null);
        }
    }

    public sealed class ReverseCommand : IAutomationCommand
    {
        public const string CommandId = "builtin.reverse";
        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "反転",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion,
                    new AutomationParameterDefinition[0]);
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            int column = request.Selection.Left;
            List<AutomationCell> populatedCells = new List<AutomationCell>();
            foreach (AutomationCell cell in request.Cells)
                if (cell.Column == column && cell.Value.Length != 0)
                    populatedCells.Add(cell);
            populatedCells.Sort(delegate(AutomationCell left, AutomationCell right)
            {
                return left.Row.CompareTo(right.Row);
            });

            List<AutomationChange> changes = new List<AutomationChange>();
            for (int index = 0; index < populatedCells.Count; index++)
                changes.Add(new AutomationChange(populatedCells[index].Row, column,
                    populatedCells[populatedCells.Count - index - 1].Value));
            return AutomationResult.Success(changes, null);
        }
    }

    public sealed class ArithmeticCommand : IAutomationCommand
    {
        public const string CommandId = "builtin.arithmetic";
        public const string OperatorParameter = "operator";
        public const string OperandParameter = "operand";

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "四則演算",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion,
                    new[]
                    {
                        new AutomationParameterDefinition(OperatorParameter, "演算子",
                            AutomationParameterType.Choice, "+", true, null, null,
                            new[] { "+", "-", "*", "/" }),
                        new AutomationParameterDefinition(OperandParameter, "値",
                            AutomationParameterType.Int32, "0", true, null, null, null)
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            string operation;
            string operandText;
            int operand;
            if (!request.Parameters.TryGetValue(OperatorParameter, out operation) ||
                (operation != "+" && operation != "-" && operation != "*" && operation != "/"))
                return AutomationResult.Failure("１文字目には\"+-*/\"記号のいずれか１文字の入力が必要");
            if (!request.Parameters.TryGetValue(OperandParameter, out operandText) ||
                !Int32.TryParse(operandText, out operand))
                return AutomationResult.Failure("入力された値を数値に変換できませんでした.");
            List<AutomationCell> cells = new List<AutomationCell>(request.Cells);
            cells.Sort(delegate(AutomationCell left, AutomationCell right)
            {
                int columnOrder = left.Column.CompareTo(right.Column);
                return columnOrder != 0 ? columnOrder : left.Row.CompareTo(right.Row);
            });

            List<AutomationChange> changes = new List<AutomationChange>();
            foreach (AutomationCell cell in cells)
            {
                if (cell.Value.Length == 0 || cell.Value == request.EmptyCellValue) continue;
                int current;
                if (!Int32.TryParse(cell.Value, out current))
                    return AutomationResult.Failure("セルの値を数値に変換できませんでした.");
                if ((operation == "*" || operation == "/") && current == 0) continue;
                if (operation == "/" && operand == 0)
                    return AutomationResult.Failure("0で除算することはできません.");

                int calculated;
                if (operation == "+") calculated = unchecked(current + operand);
                else if (operation == "-") calculated = unchecked(current - operand);
                else if (operation == "*") calculated = unchecked(current * operand);
                else calculated = current / operand;
                changes.Add(new AutomationChange(cell.Row, cell.Column,
                    calculated.ToString(CultureInfo.CurrentCulture)));
            }
            return AutomationResult.Success(changes, null);
        }
    }

    public sealed class SequentialNumberCommand : IAutomationCommand
    {
        public const string CommandId = "builtin.sequential-number";
        public const string StartParameter = "start";
        public const string StepParameter = "step";
        public const string SkipParameter = "skip";

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "連番作成",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion,
                    new[]
                    {
                        IntParameter(StartParameter, "開始番号", "1"),
                        IntParameter(StepParameter, "ステップ数", "1"),
                        new AutomationParameterDefinition(SkipParameter, "番号を飛ばす",
                            AutomationParameterType.Boolean, "false", true, null, null, null)
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            int start;
            int step;
            bool skip;
            string value;
            if (!request.Parameters.TryGetValue(StartParameter, out value) || !Int32.TryParse(value, out start) ||
                !request.Parameters.TryGetValue(StepParameter, out value) || !Int32.TryParse(value, out step) ||
                !request.Parameters.TryGetValue(SkipParameter, out value) || !Boolean.TryParse(value, out skip))
                return AutomationResult.Failure("入力された値を数値に変換できませんでした.");
            if (step == 0) return AutomationResult.Failure("ステップ数には0以外を指定してください.");

            int interval = Math.Abs(step);
            int numberStep = skip ? step : step / interval;
            int number = start;
            int column = request.Selection.Left;
            List<AutomationChange> changes = new List<AutomationChange>();
            for (int offset = 0; offset < request.Selection.RowCount; offset += interval)
            {
                changes.Add(new AutomationChange(request.Selection.Top + offset, column,
                    number.ToString(CultureInfo.CurrentCulture)));
                number = unchecked(number + numberStep);
            }
            return AutomationResult.Success(changes, null);
        }

        private static AutomationParameterDefinition IntParameter(string id, string name, string defaultValue)
        {
            return new AutomationParameterDefinition(id, name, AutomationParameterType.Int32,
                defaultValue, true, null, null, null);
        }
    }

    public sealed class RepeatNumberCommand : IAutomationCommand
    {
        public const string CommandId = "builtin.repeat-number";
        public const string StartParameter = "start";
        public const string EndParameter = "end";
        public const string RowIntervalParameter = "row_interval";
        public const string LoopParameter = "loop";
        public const string SkipParameter = "skip";
        public const string InsertParameter = "insert";

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "繰り返し",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion,
                    new[]
                    {
                        IntParameter(StartParameter, "開始番号", "1"),
                        IntParameter(EndParameter, "終了番号", "1"),
                        IntParameter(RowIntervalParameter, "行間隔", "1"),
                        IntParameter(LoopParameter, "ループ", "1"),
                        IntParameter(SkipParameter, "スキップ", "0"),
                        new AutomationParameterDefinition(InsertParameter, "挿入値",
                            AutomationParameterType.String, String.Empty, false, null, null, null)
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            int start;
            int end;
            int interval;
            int loop;
            int userSkip;
            string insert;
            if (!TryInt(request, StartParameter, out start) || !TryInt(request, EndParameter, out end) ||
                !TryInt(request, RowIntervalParameter, out interval) || !TryInt(request, LoopParameter, out loop) ||
                !TryInt(request, SkipParameter, out userSkip))
                return AutomationResult.Failure("入力された値を数値に変換できませんでした.");
            request.Parameters.TryGetValue(InsertParameter, out insert);
            insert = insert ?? String.Empty;
            if (interval <= 0) return AutomationResult.Failure("行間隔には1以上を指定してください.");
            if (loop <= 0) return AutomationResult.Failure("ループには1以上を指定してください.");
            if (end < start) return AutomationResult.Failure("終了番号には開始番号以上を指定してください.");
            if (userSkip < 0) return AutomationResult.Failure("スキップには0以上を指定してください.");

            int skip = userSkip + 1;
            int count = end - start + 1;
            if (insert.Length != 0) count *= 2;
            if (skip > 1)
            {
                count = count / skip + 1;
                if (insert.Length != 0 && count % 2 == 1) count++;
            }
            count *= loop;

            // Clear and generated writes are folded into one unique change per coordinate.
            // This preserves the single Undo group while satisfying the host's duplicate guard.
            Dictionary<string, AutomationChange> changes = new Dictionary<string, AutomationChange>();
            foreach (AutomationCell cell in request.Cells)
                changes[Key(cell.Row, cell.Column)] = new AutomationChange(cell.Row, cell.Column, String.Empty);

            int number = start;
            for (int column = request.Selection.Left;
                column < request.Selection.Left + request.Selection.ColumnCount; column++)
                for (int index = 0; index < count; index++)
                {
                    int row = request.Selection.Top + index * interval;
                    if (row >= request.RowCount)
                        return AutomationResult.Failure("繰り返し入力の書き込み先がシート範囲を超えています.");
                    string output;
                    if (insert.Length == 0 || index % 2 == 0)
                    {
                        output = number.ToString(CultureInfo.CurrentCulture);
                        number += skip;
                        if (number > end) number = start;
                    }
                    else output = insert;
                    changes[Key(row, column)] = new AutomationChange(row, column, output);
                }

            List<AutomationChange> ordered = new List<AutomationChange>(changes.Values);
            ordered.Sort(delegate(AutomationChange left, AutomationChange right)
            {
                int columnOrder = left.Column.CompareTo(right.Column);
                return columnOrder != 0 ? columnOrder : left.Row.CompareTo(right.Row);
            });
            return AutomationResult.Success(ordered, null);
        }

        private static bool TryInt(AutomationRequest request, string id, out int value)
        {
            string text;
            value = 0;
            if (!request.Parameters.TryGetValue(id, out text)) return false;
            return Int32.TryParse(text, out value);
        }

        private static string Key(int row, int column)
        {
            return row.ToString(CultureInfo.InvariantCulture) + ":" +
                column.ToString(CultureInfo.InvariantCulture);
        }

        private static AutomationParameterDefinition IntParameter(string id, string name, string defaultValue)
        {
            return new AutomationParameterDefinition(id, name, AutomationParameterType.Int32,
                defaultValue, true, null, null, null);
        }
    }
}
