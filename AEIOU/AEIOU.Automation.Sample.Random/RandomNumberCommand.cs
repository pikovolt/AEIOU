using System;
using System.Collections.Generic;
using System.Globalization;
using AEIOU.Automation;

namespace AEIOU.Automation.Sample.Random
{
    /// <summary>Example external command that fills every selected cell with a random integer.</summary>
    public sealed class RandomNumberCommand : IAutomationCommand
    {
        public const string CommandId = "sample.random-number";
        public const string MinimumParameter = "minimum";
        public const string MaximumParameter = "maximum";
        public const string StepParameter = "step";
        private readonly System.Random random;

        public RandomNumberCommand() : this(new System.Random()) { }

        internal RandomNumberCommand(System.Random random)
        {
            if (random == null) throw new ArgumentNullException("random");
            this.random = random;
        }

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "ランダム整数",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion, new[]
                    {
                        new AutomationParameterDefinition(MinimumParameter, "最小値", AutomationParameterType.Int32,
                            "1", true, Int32.MinValue, Int32.MaxValue, null),
                        new AutomationParameterDefinition(MaximumParameter, "最大値", AutomationParameterType.Int32,
                            "100", true, Int32.MinValue, Int32.MaxValue, null),
                        new AutomationParameterDefinition(StepParameter, "ステップ数", AutomationParameterType.Int32,
                            "1", true, 1, Int32.MaxValue, null)
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            if (request == null) throw new ArgumentNullException("request");
            int minimum;
            int maximum;
            int step;
            string parameterValue;
            if (!request.Parameters.TryGetValue(MinimumParameter, out parameterValue) ||
                !Int32.TryParse(parameterValue, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out minimum) ||
                !request.Parameters.TryGetValue(MaximumParameter, out parameterValue) ||
                !Int32.TryParse(parameterValue, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out maximum) ||
                !request.Parameters.TryGetValue(StepParameter, out parameterValue) ||
                !Int32.TryParse(parameterValue, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out step))
                return AutomationResult.Failure("最小値、最大値、ステップ数には整数を指定してください。");
            if (minimum > maximum)
                return AutomationResult.Failure("最小値は最大値以下にしてください。");
            if (step <= 0)
                return AutomationResult.Failure("ステップ数には1以上を指定してください。");

            List<AutomationChange> changes = new List<AutomationChange>();
            AutomationSelection selection = request.Selection;
            for (int column = selection.Left; column < selection.Left + selection.ColumnCount; column++)
                for (int row = selection.Top; row < selection.Top + selection.RowCount; row += step)
                {
                    long range = (long)maximum - minimum + 1L;
                    int value = (int)(minimum + (long)(random.NextDouble() * range));
                    changes.Add(new AutomationChange(row, column,
                        value.ToString(CultureInfo.InvariantCulture)));
                }
            return AutomationResult.Success(changes, null);
        }
    }
}
