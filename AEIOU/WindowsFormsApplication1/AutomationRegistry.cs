using System;
using System.Collections.Generic;
using global::AEIOU.Automation;

namespace AEIOU
{
    /// <summary>Stores automation commands by their stable descriptor ID.</summary>
    public sealed class AutomationRegistry
    {
        private readonly Dictionary<string, IAutomationCommand> commands =
            new Dictionary<string, IAutomationCommand>(StringComparer.Ordinal);
        private readonly Dictionary<string, AutomationCommandDescriptor> descriptors =
            new Dictionary<string, AutomationCommandDescriptor>(StringComparer.Ordinal);

        public bool TryRegister(IAutomationCommand command, out string error)
        {
            AutomationCommandDescriptor descriptor;
            return TryRegister(command, out descriptor, out error);
        }

        /// <summary>Registers a command and returns the descriptor snapshot used for registration.</summary>
        public bool TryRegister(IAutomationCommand command, out AutomationCommandDescriptor descriptor,
            out string error)
        {
            if (command == null) throw new ArgumentNullException("command");
            descriptor = command.Descriptor;
            if (descriptor == null)
            {
                error = "The command has no descriptor.";
                return false;
            }
            if (descriptor.ContractMajorVersion != AutomationContract.MajorVersion ||
                descriptor.ContractMinorVersion > AutomationContract.MinorVersion)
            {
                error = "Automation command '" + descriptor.Id + "' uses unsupported contract version " +
                    descriptor.ContractMajorVersion + "." + descriptor.ContractMinorVersion + ".";
                return false;
            }
            if (commands.ContainsKey(descriptor.Id))
            {
                error = "Automation command ID '" + descriptor.Id + "' is already registered.";
                return false;
            }
            commands.Add(descriptor.Id, command);
            descriptors.Add(descriptor.Id, descriptor);
            error = null;
            return true;
        }

        public bool TryGet(string commandId, out IAutomationCommand command)
        {
            if (commandId == null) throw new ArgumentNullException("commandId");
            return commands.TryGetValue(commandId, out command);
        }

        public bool TryGetDescriptor(string commandId, out AutomationCommandDescriptor descriptor)
        {
            if (commandId == null) throw new ArgumentNullException("commandId");
            return descriptors.TryGetValue(commandId, out descriptor);
        }
    }

    public static class BuiltInAutomationRegistry
    {
        public static AutomationRegistry Create()
        {
            AutomationRegistry registry = new AutomationRegistry();
            Register(registry, new ReplaceCommand());
            Register(registry, new ReverseCommand());
            Register(registry, new ArithmeticCommand());
            Register(registry, new SequentialNumberCommand());
            Register(registry, new RepeatNumberCommand());
            return registry;
        }

        private static void Register(AutomationRegistry registry, IAutomationCommand command)
        {
            string error;
            if (!registry.TryRegister(command, out error)) throw new InvalidOperationException(error);
        }
    }
}
