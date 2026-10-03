using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace AEIOU.Automation
{
    public static class AutomationContract
    {
        public const int MajorVersion = 1;
        public const int MinorVersion = 0;
    }

    public interface IAutomationCommand
    {
        AutomationCommandDescriptor Descriptor { get; }
        AutomationResult Execute(AutomationRequest request);
    }

    public enum AutomationParameterType
    {
        String,
        Int32,
        Boolean,
        Choice
    }

    public sealed class AutomationCommandDescriptor
    {
        private readonly ReadOnlyCollection<AutomationParameterDefinition> parameters;

        public AutomationCommandDescriptor(string id, string displayName,
            int contractMajorVersion, int contractMinorVersion,
            IEnumerable<AutomationParameterDefinition> parameters)
        {
            ValidateId(id, "id");
            if (String.IsNullOrEmpty(displayName)) throw new ArgumentException("A display name is required.", "displayName");
            if (contractMajorVersion <= 0) throw new ArgumentOutOfRangeException("contractMajorVersion");
            if (contractMinorVersion < 0) throw new ArgumentOutOfRangeException("contractMinorVersion");
            Id = id;
            DisplayName = displayName;
            ContractMajorVersion = contractMajorVersion;
            ContractMinorVersion = contractMinorVersion;
            this.parameters = Copy(parameters, "parameters").AsReadOnly();
        }

        public string Id { get; private set; }
        public string DisplayName { get; private set; }
        public int ContractMajorVersion { get; private set; }
        public int ContractMinorVersion { get; private set; }
        public ReadOnlyCollection<AutomationParameterDefinition> Parameters { get { return parameters; } }

        private static List<T> Copy<T>(IEnumerable<T> source, string name)
        {
            if (source == null) throw new ArgumentNullException(name);
            List<T> copy = new List<T>(source);
            if (copy.Contains(default(T))) throw new ArgumentException("The collection cannot contain null.", name);
            return copy;
        }

        internal static void ValidateId(string id, string parameterName)
        {
            if (String.IsNullOrEmpty(id)) throw new ArgumentException("An ID is required.", parameterName);
            for (int index = 0; index < id.Length; index++)
            {
                char character = id[index];
                bool valid = character >= 'a' && character <= 'z' ||
                    character >= '0' && character <= '9' || character == '.' || character == '_' || character == '-';
                if (!valid) throw new ArgumentException("IDs may contain only lower-case ASCII letters, digits, '.', '_' and '-'.", parameterName);
            }
        }
    }

    public sealed class AutomationParameterDefinition
    {
        private readonly ReadOnlyCollection<string> choices;

        public AutomationParameterDefinition(string id, string displayName, AutomationParameterType type,
            string defaultValue, bool isRequired, int? minimum, int? maximum, IEnumerable<string> choices)
        {
            AutomationCommandDescriptor.ValidateId(id, "id");
            if (String.IsNullOrEmpty(displayName)) throw new ArgumentException("A display name is required.", "displayName");
            if (minimum.HasValue && maximum.HasValue && minimum.Value > maximum.Value)
                throw new ArgumentException("Minimum cannot exceed maximum.");
            List<string> choiceCopy = choices == null ? new List<string>() : new List<string>(choices);
            if (choiceCopy.Contains(null)) throw new ArgumentException("Choices cannot contain null.", "choices");
            if (type == AutomationParameterType.Choice && choiceCopy.Count == 0)
                throw new ArgumentException("A choice parameter requires at least one choice.", "choices");

            Id = id;
            DisplayName = displayName;
            Type = type;
            DefaultValue = defaultValue;
            IsRequired = isRequired;
            Minimum = minimum;
            Maximum = maximum;
            this.choices = choiceCopy.AsReadOnly();
        }

        public string Id { get; private set; }
        public string DisplayName { get; private set; }
        public AutomationParameterType Type { get; private set; }
        public string DefaultValue { get; private set; }
        public bool IsRequired { get; private set; }
        public int? Minimum { get; private set; }
        public int? Maximum { get; private set; }
        public ReadOnlyCollection<string> Choices { get { return choices; } }
    }

    public sealed class AutomationSelection
    {
        public AutomationSelection(int top, int left, int rowCount, int columnCount)
        {
            if (top < 0) throw new ArgumentOutOfRangeException("top");
            if (left < 0) throw new ArgumentOutOfRangeException("left");
            if (rowCount <= 0) throw new ArgumentOutOfRangeException("rowCount");
            if (columnCount <= 0) throw new ArgumentOutOfRangeException("columnCount");
            Top = top;
            Left = left;
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        public int Top { get; private set; }
        public int Left { get; private set; }
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
    }

    public sealed class AutomationCell
    {
        public AutomationCell(int row, int column, string value)
        {
            if (row < 0) throw new ArgumentOutOfRangeException("row");
            if (column < 0) throw new ArgumentOutOfRangeException("column");
            if (value == null) throw new ArgumentNullException("value");
            Row = row;
            Column = column;
            Value = value;
        }

        public int Row { get; private set; }
        public int Column { get; private set; }
        public string Value { get; private set; }
    }

    public sealed class AutomationRequest
    {
        private readonly ReadOnlyCollection<AutomationCell> cells;
        private readonly AutomationParameterValues parameters;

        public AutomationRequest(int rowCount, int columnCount, long sheetGeneration,
            AutomationSelection selection, IEnumerable<AutomationCell> cells,
            IDictionary<string, string> parameters, string emptyCellValue)
        {
            if (rowCount <= 0) throw new ArgumentOutOfRangeException("rowCount");
            if (columnCount <= 0) throw new ArgumentOutOfRangeException("columnCount");
            if (selection == null) throw new ArgumentNullException("selection");
            if (cells == null) throw new ArgumentNullException("cells");
            if (parameters == null) throw new ArgumentNullException("parameters");
            if (emptyCellValue == null) throw new ArgumentNullException("emptyCellValue");
            if (selection.Top > rowCount - selection.RowCount || selection.Left > columnCount - selection.ColumnCount)
                throw new ArgumentException("The selection must be inside the sheet.", "selection");
            List<AutomationCell> cellCopy = new List<AutomationCell>(cells);
            if (cellCopy.Contains(null)) throw new ArgumentException("Cells cannot contain null.", "cells");
            foreach (AutomationCell cell in cellCopy)
            {
                if (cell.Row < selection.Top || cell.Row >= selection.Top + selection.RowCount ||
                    cell.Column < selection.Left || cell.Column >= selection.Left + selection.ColumnCount)
                    throw new ArgumentException("Snapshot cells must be inside the selection.", "cells");
            }
            Dictionary<string, string> parameterCopy = new Dictionary<string, string>(parameters, StringComparer.Ordinal);
            foreach (KeyValuePair<string, string> item in parameterCopy)
                if (item.Key == null || item.Value == null) throw new ArgumentException("Parameters cannot contain null.", "parameters");

            RowCount = rowCount;
            ColumnCount = columnCount;
            SheetGeneration = sheetGeneration;
            Selection = selection;
            this.cells = cellCopy.AsReadOnly();
            this.parameters = new AutomationParameterValues(parameterCopy);
            EmptyCellValue = emptyCellValue;
        }

        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
        public long SheetGeneration { get; private set; }
        public AutomationSelection Selection { get; private set; }
        public ReadOnlyCollection<AutomationCell> Cells { get { return cells; } }
        public AutomationParameterValues Parameters { get { return parameters; } }
        public string EmptyCellValue { get; private set; }
    }

    public sealed class AutomationParameterValues : IEnumerable<KeyValuePair<string, string>>
    {
        private readonly Dictionary<string, string> values;

        internal AutomationParameterValues(IDictionary<string, string> values)
        {
            this.values = new Dictionary<string, string>(values, StringComparer.Ordinal);
        }

        public int Count { get { return values.Count; } }
        public string this[string id] { get { return values[id]; } }
        public bool Contains(string id) { return values.ContainsKey(id); }
        public bool TryGetValue(string id, out string value) { return values.TryGetValue(id, out value); }
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() { return values.GetEnumerator(); }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return GetEnumerator(); }
    }

    public sealed class AutomationChange
    {
        public AutomationChange(int row, int column, string value)
        {
            Row = row;
            Column = column;
            Value = value;
        }

        public int Row { get; private set; }
        public int Column { get; private set; }
        public string Value { get; private set; }
    }

    public sealed class AutomationResult
    {
        private readonly ReadOnlyCollection<AutomationChange> changes;

        private AutomationResult(bool succeeded, IEnumerable<AutomationChange> changes, string error, string message)
        {
            List<AutomationChange> copy = changes == null ? new List<AutomationChange>() : new List<AutomationChange>(changes);
            Succeeded = succeeded;
            this.changes = copy.AsReadOnly();
            Error = error;
            Message = message;
        }

        public bool Succeeded { get; private set; }
        public ReadOnlyCollection<AutomationChange> Changes { get { return changes; } }
        public string Error { get; private set; }
        public string Message { get; private set; }

        public static AutomationResult Success(IEnumerable<AutomationChange> changes, string message)
        {
            if (changes == null) throw new ArgumentNullException("changes");
            return new AutomationResult(true, changes, null, message);
        }

        public static AutomationResult Failure(string error)
        {
            if (String.IsNullOrEmpty(error)) throw new ArgumentException("An error is required.", "error");
            return new AutomationResult(false, null, error, null);
        }
    }
}
