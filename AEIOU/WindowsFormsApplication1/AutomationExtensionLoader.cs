using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using AEIOU.Automation;

namespace AEIOU
{
    public sealed class AutomationExtensionLoadResult
    {
        private readonly List<IAutomationCommand> commands = new List<IAutomationCommand>();
        private readonly List<AutomationCommandDescriptor> descriptors =
            new List<AutomationCommandDescriptor>();
        private readonly List<string> diagnostics = new List<string>();

        public IList<IAutomationCommand> Commands { get { return commands.AsReadOnly(); } }
        public IList<AutomationCommandDescriptor> Descriptors { get { return descriptors.AsReadOnly(); } }
        public IList<string> Diagnostics { get { return diagnostics.AsReadOnly(); } }
        internal void AddCommand(IAutomationCommand command, AutomationCommandDescriptor descriptor)
        {
            commands.Add(command);
            descriptors.Add(descriptor);
        }
        internal void AddDiagnostic(string diagnostic) { diagnostics.Add(diagnostic); }
    }

    /// <summary>Discovers trusted automation commands from one non-recursive directory.</summary>
    public sealed class AutomationExtensionLoader
    {
        public AutomationExtensionLoadResult Load(string directory, AutomationRegistry registry)
        {
            if (directory == null) throw new ArgumentNullException("directory");
            if (registry == null) throw new ArgumentNullException("registry");

            AutomationExtensionLoadResult result = new AutomationExtensionLoadResult();
            if (!Directory.Exists(directory)) return result;

            string[] paths;
            try { paths = Directory.GetFiles(directory, "*.dll", SearchOption.TopDirectoryOnly); }
            catch (Exception exception)
            {
                Record(result, directory, null, exception);
                return result;
            }
            Array.Sort(paths, StringComparer.OrdinalIgnoreCase);
            foreach (string path in paths) LoadAssembly(path, registry, result);
            return result;
        }

        private static void LoadAssembly(string path, AutomationRegistry registry,
            AutomationExtensionLoadResult result)
        {
            Assembly assembly;
            try { assembly = Assembly.LoadFrom(path); }
            catch (Exception exception)
            {
                Record(result, path, null, exception);
                return;
            }

            Type[] types;
            try { types = assembly.GetTypes(); }
            catch (ReflectionTypeLoadException exception)
            {
                types = exception.Types;
                foreach (Exception loaderException in exception.LoaderExceptions)
                    if (loaderException != null) Record(result, path, null, loaderException);
            }
            catch (Exception exception)
            {
                Record(result, path, null, exception);
                return;
            }

            foreach (Type type in types)
            {
                if (type == null || !type.IsPublic || type.IsAbstract || type.IsInterface ||
                    !typeof(IAutomationCommand).IsAssignableFrom(type) ||
                    type.GetConstructor(Type.EmptyTypes) == null) continue;
                IAutomationCommand command;
                try { command = (IAutomationCommand)Activator.CreateInstance(type); }
                catch (Exception exception)
                {
                    Record(result, path, type.FullName, Unwrap(exception));
                    continue;
                }

                try
                {
                    AutomationCommandDescriptor descriptor;
                    string error;
                    if (!registry.TryRegister(command, out descriptor, out error))
                    {
                        result.AddDiagnostic(Format(path, descriptor == null ? null : descriptor.Id, error));
                        continue;
                    }
                    result.AddCommand(command, descriptor);
                }
                catch (Exception exception)
                {
                    Record(result, path, type.FullName, exception);
                }
            }
        }

        public static void AppendDiagnostics(string logPath, IEnumerable<string> diagnostics)
        {
            if (String.IsNullOrEmpty(logPath) || diagnostics == null) return;
            try
            {
                using (StreamWriter writer = new StreamWriter(logPath, true))
                    foreach (string diagnostic in diagnostics)
                        writer.WriteLine(DateTime.Now.ToString("s") + " " + diagnostic);
            }
            catch (Exception) { /* Logging must never prevent application startup. */ }
        }

        private static Exception Unwrap(Exception exception)
        {
            TargetInvocationException invocation = exception as TargetInvocationException;
            return invocation != null && invocation.InnerException != null ? invocation.InnerException : exception;
        }

        private static void Record(AutomationExtensionLoadResult result, string path, string commandId,
            Exception exception)
        {
            result.AddDiagnostic(Format(path, commandId, exception.GetType().Name + ": " + exception.Message));
        }

        private static string Format(string path, string commandId, string message)
        {
            return "DLL='" + path + "' command='" + (commandId ?? "unknown") + "' " + message;
        }
    }
}
