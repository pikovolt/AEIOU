using System;
using System.Collections.Generic;
using System.Globalization;
using global::AEIOU.Automation;

namespace AEIOU
{
    /// <summary>Executes untrusted command output only after validating the complete change set.</summary>
    public sealed class AutomationHost
    {
        private readonly int maximumChangeCount;
        private readonly AutomationRegistry registry;

        public AutomationHost(int maximumChangeCount)
            : this(maximumChangeCount, null)
        {
        }

        public AutomationHost(int maximumChangeCount, AutomationRegistry registry)
        {
            if (maximumChangeCount <= 0) throw new ArgumentOutOfRangeException("maximumChangeCount");
            this.maximumChangeCount = maximumChangeCount;
            this.registry = registry;
        }

        public AutomationHostResult Execute(string commandId, AutomationRequest request,
            IAutomationChangeTarget target)
        {
            if (String.IsNullOrEmpty(commandId))
                return AutomationHostResult.Rejected("An automation command ID is required.");
            IAutomationCommand command;
            if (registry == null || !registry.TryGet(commandId, out command))
                return AutomationHostResult.Rejected("Automation command '" + commandId + "' is not registered.");
            AutomationCommandDescriptor descriptor;
            if (!registry.TryGetDescriptor(commandId, out descriptor))
                return AutomationHostResult.Rejected("Automation command '" + commandId + "' has no descriptor.");
            return Execute(command, descriptor, request, target);
        }

        public AutomationHostResult Execute(IAutomationCommand command, AutomationRequest request,
            IAutomationChangeTarget target)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (request == null) throw new ArgumentNullException("request");
            if (target == null) throw new ArgumentNullException("target");

            AutomationCommandDescriptor descriptor;
            try { descriptor = command.Descriptor; }
            catch (Exception exception)
            {
                return AutomationHostResult.Faulted("unknown", exception);
            }
            return Execute(command, descriptor, request, target);
        }

        private AutomationHostResult Execute(IAutomationCommand command,
            AutomationCommandDescriptor descriptor, AutomationRequest request, IAutomationChangeTarget target)
        {
            if (descriptor == null)
                return AutomationHostResult.Rejected("The command has no descriptor.");
            if (descriptor.ContractMajorVersion != AutomationContract.MajorVersion ||
                descriptor.ContractMinorVersion > AutomationContract.MinorVersion)
                return AutomationHostResult.Rejected("The command contract version is not supported.");
            string parameterError = AutomationParameterAdapter.Validate(descriptor, request.Parameters);
            if (parameterError != null) return AutomationHostResult.Rejected(parameterError);

            AutomationResult result;
            try
            {
                result = command.Execute(request);
            }
            catch (Exception exception)
            {
                return AutomationHostResult.Faulted(descriptor.Id, exception);
            }

            string validationError = Validate(result, request);
            if (validationError != null)
                return AutomationHostResult.Rejected(validationError);
            if (!result.Succeeded)
                return AutomationHostResult.Rejected(result.Error);
            if (result.Changes.Count == 0)
                return AutomationHostResult.Applied(0, result.Message);

            // The target performs the generation comparison and write as one operation so that
            // the sheet cannot change between the final check and the write group.
            if (!target.TryApply(request.SheetGeneration, result.Changes, descriptor.DisplayName))
                return AutomationHostResult.Rejected("The sheet changed while the command was running.");
            return AutomationHostResult.Applied(result.Changes.Count, result.Message);
        }

        private string Validate(AutomationResult result, AutomationRequest request)
        {
            if (result == null) return "The command returned no result.";
            if (!result.Succeeded)
                return String.IsNullOrEmpty(result.Error) ? "The command failed without an error message." : null;
            if (result.Changes.Count > maximumChangeCount)
                return "The command returned too many changes.";

            HashSet<string> coordinates = new HashSet<string>(StringComparer.Ordinal);
            foreach (AutomationChange change in result.Changes)
            {
                if (change == null) return "The change set contains a null change.";
                if (change.Row < 0 || change.Row >= request.RowCount ||
                    change.Column < 0 || change.Column >= request.ColumnCount)
                    return "The change set contains an out-of-range cell.";
                if (change.Value == null) return "The change set contains a null value.";
                string coordinate = change.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                    change.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!coordinates.Add(coordinate)) return "The change set contains a duplicate cell.";
            }
            return null;
        }
    }

    internal static class AutomationParameterAdapter
    {
        public static string Validate(AutomationCommandDescriptor descriptor, AutomationParameterValues values)
        {
            Dictionary<string, AutomationParameterDefinition> definitions =
                new Dictionary<string, AutomationParameterDefinition>(StringComparer.Ordinal);
            foreach (AutomationParameterDefinition definition in descriptor.Parameters)
            {
                if (definitions.ContainsKey(definition.Id))
                    return "The command descriptor contains a duplicate parameter ID.";
                definitions.Add(definition.Id, definition);
            }
            foreach (KeyValuePair<string, string> value in values)
                if (!definitions.ContainsKey(value.Key))
                    return "Parameter '" + value.Key + "' is not defined for this command.";

            foreach (AutomationParameterDefinition definition in descriptor.Parameters)
            {
                string value;
                bool booleanValue;
                if (!values.TryGetValue(definition.Id, out value))
                {
                    if (definition.IsRequired) return "Required parameter '" + definition.Id + "' is missing.";
                    continue;
                }
                if (definition.IsRequired && value.Length == 0)
                    return "Required parameter '" + definition.Id + "' is empty.";

                int number;
                if (definition.Type == AutomationParameterType.Int32)
                {
                    if (!Int32.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out number))
                        return "Parameter '" + definition.Id + "' is not a valid Int32 value.";
                    if (definition.Minimum.HasValue && number < definition.Minimum.Value ||
                        definition.Maximum.HasValue && number > definition.Maximum.Value)
                        return "Parameter '" + definition.Id + "' is outside its allowed range.";
                }
                else if (definition.Type == AutomationParameterType.Boolean && !Boolean.TryParse(value, out booleanValue))
                    return "Parameter '" + definition.Id + "' is not a valid Boolean value.";
                else if (definition.Type == AutomationParameterType.Choice && !definition.Choices.Contains(value))
                    return "Parameter '" + definition.Id + "' is not an allowed choice.";
            }
            return null;
        }

    }

    public interface IAutomationChangeTarget
    {
        bool TryApply(long expectedGeneration, IList<AutomationChange> changes, string operationName);
    }

    /// <summary>Owns the last result of a modeless automation dialog and replaces only that result.</summary>
    public sealed class AutomationSession
    {
        private readonly AutomationHost host;
        private object lastApplication;

        public AutomationSession(AutomationHost host)
        {
            if (host == null) throw new ArgumentNullException("host");
            this.host = host;
        }

        public AutomationHostResult Execute(IAutomationCommand command, AutomationRequest request,
            IAutomationSessionTarget target)
        {
            if (target == null) throw new ArgumentNullException("target");
            SessionChangeTarget changeTarget = new SessionChangeTarget(target, lastApplication);
            AutomationHostResult result = host.Execute(command, request, changeTarget);
            if (result.Succeeded && changeTarget.Applied)
                lastApplication = changeTarget.Application;
            return result;
        }

        public AutomationHostResult Execute(string commandId, AutomationRequest request,
            IAutomationSessionTarget target)
        {
            if (target == null) throw new ArgumentNullException("target");
            SessionChangeTarget changeTarget = new SessionChangeTarget(target, lastApplication);
            AutomationHostResult result = host.Execute(commandId, request, changeTarget);
            if (result.Succeeded && changeTarget.Applied)
                lastApplication = changeTarget.Application;
            return result;
        }

        public void Close()
        {
            lastApplication = null;
        }

        private sealed class SessionChangeTarget : IAutomationChangeTarget
        {
            private readonly IAutomationSessionTarget target;
            private readonly object previousApplication;

            public SessionChangeTarget(IAutomationSessionTarget target, object previousApplication)
            {
                this.target = target;
                this.previousApplication = previousApplication;
            }

            public bool Applied { get; private set; }
            public object Application { get; private set; }

            public bool TryApply(long expectedGeneration, IList<AutomationChange> changes, string operationName)
            {
                object application;
                if (!target.TryReplace(expectedGeneration, previousApplication, changes,
                    operationName, out application)) return false;
                Applied = true;
                Application = application;
                return true;
            }
        }
    }

    public interface IAutomationSessionTarget
    {
        bool TryReplace(long expectedGeneration, object previousApplication,
            IList<AutomationChange> changes, string operationName, out object application);
    }

    public sealed class AutomationHostResult
    {
        private AutomationHostResult(bool succeeded, int appliedChangeCount, string error, string message, Exception exception)
        {
            Succeeded = succeeded;
            AppliedChangeCount = appliedChangeCount;
            Error = error;
            Message = message;
            Exception = exception;
        }

        public bool Succeeded { get; private set; }
        public int AppliedChangeCount { get; private set; }
        public string Error { get; private set; }
        public string Message { get; private set; }
        public Exception Exception { get; private set; }

        internal static AutomationHostResult Applied(int count, string message)
        {
            return new AutomationHostResult(true, count, null, message, null);
        }

        internal static AutomationHostResult Rejected(string error)
        {
            return new AutomationHostResult(false, 0, error, null, null);
        }

        internal static AutomationHostResult Faulted(string commandId, Exception exception)
        {
            return new AutomationHostResult(false, 0,
                "Automation command '" + commandId + "' failed.", null, exception);
        }
    }
}
