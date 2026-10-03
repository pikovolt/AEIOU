using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using AEIOU.Automation;
using AEIOU.Automation.Sample.Random;

namespace AEIOU.Automation.Tests
{
    internal static class Program
    {
        private static int failures;

        private static int Main()
        {
            Run("row insert calculations preserve values and write order", RowInsertCalculationsPreserveBehavior);
            Run("row delete calculations preserve values and write order", RowDeleteCalculationsPreserveBehavior);
            Run("row calculations clear without shifts at the sheet end", RowCalculationsHandleZeroMoveLength);
            Run("row calculations validate all inputs before reading", RowCalculationsValidateBeforeReading);
            Run("column insert calculations preserve values metadata and order", ColumnInsertCalculationsPreserveBehavior);
            Run("column delete calculations preserve values metadata and order", ColumnDeleteCalculationsPreserveBehavior);
            Run("column calculations handle sheet edges", ColumnCalculationsHandleEdges);
            Run("column calculations validate before reading", ColumnCalculationsValidateBeforeReading);
            Run("cell write batches reject out-of-range coordinates", CellWriteBatchRejectsOutOfRangeCoordinates);
            Run("cell write batches reject duplicate coordinates", CellWriteBatchRejectsDuplicateCoordinates);
            Run("valid changes are applied once", ValidChangesAreAppliedOnce);
            Run("empty changes do not open a write group", EmptyChangesDoNotApply);
            Run("out-of-range result is rejected atomically", OutOfRangeIsRejected);
            Run("duplicate result is rejected atomically", DuplicateIsRejected);
            Run("null value is rejected atomically", NullValueIsRejected);
            Run("change limit is rejected atomically", ChangeLimitIsRejected);
            Run("generation mismatch is rejected atomically", GenerationMismatchIsRejected);
            Run("command exception is isolated", CommandExceptionIsIsolated);
            Run("built-in commands resolve through registry", BuiltInsResolveThroughRegistry);
            Run("registry rejects duplicate IDs", RegistryRejectsDuplicateIds);
            Run("registry keeps the registered descriptor snapshot", RegistryKeepsDescriptorSnapshot);
            Run("host rejects unknown IDs", HostRejectsUnknownIds);
            Run("host validates descriptor parameters", HostValidatesDescriptorParameters);
            Run("replace matches only the left column exactly", ReplaceMatchesLeftColumnExactly);
            Run("replace rejects empty values", ReplaceRejectsEmptyValues);
            Run("reverse preserves empty positions", ReversePreservesEmptyPositions);
            Run("arithmetic handles all operators and ignored cells", ArithmeticHandlesOperatorsAndIgnoredCells);
            Run("arithmetic rejects invalid cells atomically", ArithmeticRejectsInvalidCellsAtomically);
            Run("arithmetic rejects division by zero", ArithmeticRejectsDivisionByZero);
            Run("arithmetic ignores division by zero for ignored cells", ArithmeticIgnoresDivisionByZeroForIgnoredCells);
            Run("sequential number preserves step and skip behavior", SequentialNumberPreservesBehavior);
            Run("sequential number rejects zero step", SequentialNumberRejectsZeroStep);
            Run("repeat number handles insert skip loop and columns", RepeatNumberHandlesParameters);
            Run("repeat number clears selection and rejects invalid ranges", RepeatNumberClearsAndValidates);
            Run("automation session replaces only its last result", AutomationSessionReplacesOnlyItsLastResult);
            Run("extension discovery isolates failures and registers valid commands", ExtensionDiscoveryIsolatesFailures);
            Run("extension discovery rejects duplicate IDs", ExtensionDiscoveryRejectsDuplicates);
            Run("random sample respects the frame step", RandomSampleRespectsFrameStep);
            Run("AE clipboard sample writes without cell changes", AeClipboardSampleWritesWithoutCellChanges);
            Console.WriteLine(failures == 0 ? "All automation host tests passed." : failures + " test(s) failed.");
            return failures == 0 ? 0 : 1;
        }

        private static void RowInsertCalculationsPreserveBehavior()
        {
            AssertRowEditResult(true, 0, 1, new[] { "", "0", "1", "2", "3" });
            AssertRowEditResult(true, 2, 1, new[] { "0", "1", "", "2", "3" });
            AssertRowEditResult(true, 4, 1, new[] { "0", "1", "2", "3", "" });
            AssertRowEditResult(true, 1, 2, new[] { "0", "", "", "1", "2" });

            IList<CellWriteEntry> writes = CalculateRows(true, 1, 2, ValueAt);
            AssertWrites(writes, new[]
            {
                "4:0=2", "3:0=1", "4:1=12", "3:1=11",
                "1:0=", "1:1=", "2:0=", "2:1="
            });
        }

        private static void RowDeleteCalculationsPreserveBehavior()
        {
            AssertRowEditResult(false, 0, 1, new[] { "1", "2", "3", "4", "" });
            AssertRowEditResult(false, 2, 1, new[] { "0", "1", "3", "4", "" });
            AssertRowEditResult(false, 4, 1, new[] { "0", "1", "2", "3", "" });
            AssertRowEditResult(false, 1, 2, new[] { "0", "3", "4", "", "" });
            AssertRowEditResult(false, 0, 5, new[] { "", "", "", "", "" });

            IList<CellWriteEntry> writes = CalculateRows(false, 1, 2, ValueAt);
            AssertWrites(writes, new[]
            {
                "1:0=3", "2:0=4", "1:1=13", "2:1=14",
                "3:0=", "3:1=", "4:0=", "4:1="
            });
        }

        private static void RowCalculationsHandleZeroMoveLength()
        {
            int reads = 0;
            CellValueReader reader = delegate(int row, int column) { reads++; return null; };
            IList<CellWriteEntry> inserted = CalculateRows(true, 3, 2, reader);
            IList<CellWriteEntry> deleted = CalculateRows(false, 3, 2, reader);
            Assert(reads == 0, "an end edit must not read cells when there is nothing to shift");
            Assert(inserted.Count == 4 && deleted.Count == 4,
                "an end edit must still emit every required clear");
            Assert(inserted[0].Value == String.Empty && deleted[0].Value == String.Empty,
                "clear values must be normalized to empty strings");

            IList<CellWriteEntry> normalized = SheetRowEditCalculator.CreateInsertRows(
                2, 1, 0, 1, delegate { return null; });
            Assert(normalized[0].Value == String.Empty, "reader null values must be normalized");
        }

        private static void RowCalculationsValidateBeforeReading()
        {
            int reads = 0;
            CellValueReader reader = delegate(int row, int column) { reads++; return "value"; };
            AssertThrows(delegate { SheetRowEditCalculator.CreateInsertRows(0, 1, 0, 1, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateInsertRows(5, 0, 0, 1, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateDeleteRows(5, 1, -1, 1, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateDeleteRows(5, 1, 5, 1, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateInsertRows(5, 1, 0, 0, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateDeleteRows(5, 1, 4, 2, reader); });
            AssertThrows(delegate { SheetRowEditCalculator.CreateInsertRows(5, 1, 0, 1, null); });
            Assert(reads == 0, "invalid input must be rejected before the snapshot is read");
        }

        private static void ColumnInsertCalculationsPreserveBehavior()
        {
            IList<ColumnWriteEntry> writes = SheetColumnEditCalculator.CreateInsertColumn(
                2, 3, 1, ColumnValueAt, ColumnHeaderAt, ColumnUsedCountAt);
            AssertColumnWrites(writes, new[]
            {
                "3=H2;102;2:0,2:1", "2=H1;101;1:0,1:1", "1=;0;,"
            });
        }

        private static void ColumnDeleteCalculationsPreserveBehavior()
        {
            IList<ColumnWriteEntry> writes = SheetColumnEditCalculator.CreateDeleteColumn(
                2, 4, 1, ColumnValueAt, ColumnHeaderAt, ColumnUsedCountAt);
            AssertColumnWrites(writes, new[]
            {
                "1=H2;102;2:0,2:1", "2=H3;103;3:0,3:1"
            });
        }

        private static void ColumnCalculationsHandleEdges()
        {
            IList<ColumnWriteEntry> insertedFirst = SheetColumnEditCalculator.CreateInsertColumn(
                1, 2, 0, ColumnValueAt, ColumnHeaderAt, ColumnUsedCountAt);
            Assert(insertedFirst.Count == 3 && insertedFirst[0].Column == 2 &&
                insertedFirst[2].Column == 0, "first-column insertion must shift every column from the end");

            IList<ColumnWriteEntry> insertedLast = SheetColumnEditCalculator.CreateInsertColumn(
                1, 2, 1, ColumnValueAt, ColumnHeaderAt, ColumnUsedCountAt);
            Assert(insertedLast.Count == 2 && insertedLast[0].Column == 2 &&
                insertedLast[1].Column == 1, "last-column insertion must shift and clear the selected column");

            IList<ColumnWriteEntry> deletedLast = SheetColumnEditCalculator.CreateDeleteColumn(
                1, 2, 1, ColumnValueAt, ColumnHeaderAt, ColumnUsedCountAt);
            Assert(deletedLast.Count == 0, "last-column deletion needs no writes before resize");

            IList<ColumnWriteEntry> normalized = SheetColumnEditCalculator.CreateInsertColumn(
                1, 1, 0, delegate { return null; }, delegate { return null; }, delegate { return 0; });
            Assert(normalized[0].Header == String.Empty && normalized[0].GetValue(0) == String.Empty,
                "column snapshots must normalize null strings");
        }

        private static void ColumnCalculationsValidateBeforeReading()
        {
            int reads = 0;
            CellValueReader values = delegate { reads++; return "value"; };
            ColumnHeaderReader headers = delegate { reads++; return "header"; };
            ColumnUsedCountReader used = delegate { reads++; return 1; };
            AssertThrows(delegate { SheetColumnEditCalculator.CreateInsertColumn(0, 1, 0, values, headers, used); });
            AssertThrows(delegate { SheetColumnEditCalculator.CreateDeleteColumn(1, 0, 0, values, headers, used); });
            AssertThrows(delegate { SheetColumnEditCalculator.CreateInsertColumn(1, 1, 1, values, headers, used); });
            AssertThrows(delegate { SheetColumnEditCalculator.CreateDeleteColumn(1, 1, 0, null, headers, used); });
            AssertThrows(delegate { SheetColumnEditCalculator.CreateInsertColumn(1, 1, 0, values, null, used); });
            AssertThrows(delegate { SheetColumnEditCalculator.CreateDeleteColumn(1, 1, 0, values, headers, null); });
            Assert(reads == 0, "invalid column input must be rejected before the snapshot is read");
        }

        private static string ColumnValueAt(int row, int column)
        {
            return column + ":" + row;
        }

        private static string ColumnHeaderAt(int column)
        {
            return "H" + column;
        }

        private static int ColumnUsedCountAt(int column)
        {
            return 100 + column;
        }

        private static void AssertColumnWrites(IList<ColumnWriteEntry> writes, string[] expected)
        {
            Assert(writes.Count == expected.Length, "unexpected column write count");
            for (int index = 0; index < expected.Length; index++)
            {
                ColumnWriteEntry write = writes[index];
                string actual = write.Column + "=" + write.Header + ";" + write.UsedCount + ";";
                for (int row = 0; row < write.ValueCount; row++)
                {
                    if (row > 0) actual += ",";
                    actual += write.GetValue(row);
                }
                Assert(actual == expected[index], "unexpected column write at index " + index + ": " + actual);
            }
        }

        private static void CellWriteBatchRejectsOutOfRangeCoordinates()
        {
            IList<CellWriteEntry> writes = new List<CellWriteEntry>
            {
                new CellWriteEntry(0, 0, "valid"),
                new CellWriteEntry(5, 0, "invalid")
            };
            AssertThrows(delegate { CellWriteBatch.Validate(writes, 5, 2); });
        }

        private static void CellWriteBatchRejectsDuplicateCoordinates()
        {
            IList<CellWriteEntry> writes = new List<CellWriteEntry>
            {
                new CellWriteEntry(1, 1, "first"),
                new CellWriteEntry(1, 1, "second")
            };
            AssertThrows(delegate { CellWriteBatch.Validate(writes, 5, 2); });
        }

        private static IList<CellWriteEntry> CalculateRows(
            bool insert, int row, int count, CellValueReader reader)
        {
            return insert
                ? SheetRowEditCalculator.CreateInsertRows(5, 2, row, count, reader)
                : SheetRowEditCalculator.CreateDeleteRows(5, 2, row, count, reader);
        }

        private static string ValueAt(int row, int column)
        {
            return (column * 10 + row).ToString();
        }

        private static void AssertRowEditResult(bool insert, int row, int count, string[] expectedFirstColumn)
        {
            string[,] values = new string[5, 2];
            for (int currentRow = 0; currentRow < 5; currentRow++)
            {
                values[currentRow, 0] = ValueAt(currentRow, 0);
                values[currentRow, 1] = ValueAt(currentRow, 1);
            }
            IList<CellWriteEntry> writes = CalculateRows(insert, row, count,
                delegate(int sourceRow, int column) { return values[sourceRow, column]; });
            CellWriteBatch.Validate(writes, 5, 2);
            foreach (CellWriteEntry write in writes)
            {
                values[write.Row, write.Col] = write.Value;
            }
            for (int currentRow = 0; currentRow < 5; currentRow++)
            {
                Assert(values[currentRow, 0] == expectedFirstColumn[currentRow],
                    "unexpected row-edit result at row " + currentRow);
            }
        }

        private static void AssertWrites(IList<CellWriteEntry> writes, string[] expected)
        {
            Assert(writes.Count == expected.Length, "unexpected write count");
            for (int index = 0; index < expected.Length; index++)
            {
                CellWriteEntry write = writes[index];
                string actual = write.Row + ":" + write.Col + "=" + write.Value;
                Assert(actual == expected[index], "unexpected write at index " + index + ": " + actual);
            }
        }

        private static void AssertThrows(Action action)
        {
            bool threw = false;
            try { action(); }
            catch (ArgumentException) { threw = true; }
            Assert(threw, "invalid row calculation input must throw an argument exception");
        }

        private static void RandomSampleRespectsFrameStep()
        {
            AutomationRequest request = new AutomationRequest(4, 5, 1,
                new AutomationSelection(1, 2, 3, 3), new AutomationCell[0],
                new Dictionary<string, string>
                {
                    { RandomNumberCommand.MinimumParameter, "-7" },
                    { RandomNumberCommand.MaximumParameter, "-7" },
                    { RandomNumberCommand.StepParameter, "2" }
                }, String.Empty);
            RandomNumberCommand command = new RandomNumberCommand();
            AutomationResult result = command.Execute(request);

            Assert(result.Succeeded && result.Changes.Count == 6,
                "step two must produce changes for every other frame in each selected column");
            HashSet<string> coordinates = new HashSet<string>();
            foreach (AutomationChange change in result.Changes)
            {
                Assert(change.Value == "-7", "inclusive equal bounds must produce that value");
                coordinates.Add(change.Row + ":" + change.Column);
            }
            Assert(coordinates.Count == 6 && coordinates.Contains("1:2") && coordinates.Contains("1:3") &&
                coordinates.Contains("1:4") && coordinates.Contains("3:2") && coordinates.Contains("3:3") &&
                coordinates.Contains("3:4") && !coordinates.Contains("2:2"),
                "the changes must cover stepped frames in every column without duplicates");

            AutomationRequest invalid = new AutomationRequest(1, 1, 1,
                new AutomationSelection(0, 0, 1, 1), new AutomationCell[0],
                new Dictionary<string, string>
                {
                    { RandomNumberCommand.MinimumParameter, "2" },
                    { RandomNumberCommand.MaximumParameter, "1" },
                    { RandomNumberCommand.StepParameter, "1" }
                }, String.Empty);
            Assert(!command.Execute(invalid).Succeeded, "minimum greater than maximum must be rejected");

            AutomationRequest zeroStep = new AutomationRequest(1, 1, 1,
                new AutomationSelection(0, 0, 1, 1), new AutomationCell[0],
                new Dictionary<string, string>
                {
                    { RandomNumberCommand.MinimumParameter, "1" },
                    { RandomNumberCommand.MaximumParameter, "2" },
                    { RandomNumberCommand.StepParameter, "0" }
                }, String.Empty);
            Assert(!command.Execute(zeroStep).Succeeded, "zero step must be rejected");
        }

        private static void AeClipboardSampleWritesWithoutCellChanges()
        {
            FakeClipboardWriter clipboard = new FakeClipboardWriter();
            AeClipboardCommand command = new AeClipboardCommand(clipboard);
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add(AeClipboardCommand.VersionParameter, "9.0");
            parameters.Add(AeClipboardCommand.FpsParameter, "24");
            parameters.Add(AeClipboardCommand.FirstFrameParameter, "1");
            parameters.Add(AeClipboardCommand.DirectParameter, "false");
            AutomationResult result = command.Execute(CommandRequest(
                new[] { "1", "", "25" }, parameters, 3, 1));

            Assert(result.Succeeded && result.Changes.Count == 0,
                "clipboard-only commands must not return cell changes");
            Assert(clipboard.WriteCount == 1 && clipboard.Text.IndexOf("Adobe After Effects 9.0 Keyframe Data") >= 0,
                "the sample must write AE keyframe data once");
            Assert(clipboard.Text.IndexOf("\t0\t0\r\n") >= 0 && clipboard.Text.IndexOf("\t2\t1\r\n") >= 0,
                "the sample must convert selected frame values to seconds and preserve row offsets");

            AutomationResult invalid = command.Execute(CommandRequest(
                new[] { "not-a-number" }, parameters, 1, 1));
            Assert(!invalid.Succeeded && clipboard.WriteCount == 1,
                "invalid cells must be rejected before touching the clipboard");
        }

        private static void ValidChangesAreAppliedOnce()
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = Host().Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "A"), new AutomationChange(1, 1, "B") }, "done");
            }), Request(), target);
            Assert(result.Succeeded && result.AppliedChangeCount == 2, "success result expected");
            Assert(target.ApplyCount == 1 && target.Values.Count == 2, "one complete write group expected");
        }

        private static void EmptyChangesDoNotApply()
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = Host().Execute(Command(delegate { return AutomationResult.Success(new AutomationChange[0], null); }), Request(), target);
            Assert(result.Succeeded && target.ApplyCount == 0, "empty result must not touch target");
        }

        private static void OutOfRangeIsRejected()
        {
            AssertRejectedWithoutWrites(new[] { new AutomationChange(0, 0, "valid"), new AutomationChange(2, 0, "invalid") });
        }

        private static void DuplicateIsRejected()
        {
            AssertRejectedWithoutWrites(new[] { new AutomationChange(0, 0, "A"), new AutomationChange(0, 0, "B") });
        }

        private static void NullValueIsRejected()
        {
            AssertRejectedWithoutWrites(new[] { new AutomationChange(0, 0, null) });
        }

        private static void ChangeLimitIsRejected()
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = new AutomationHost(1).Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "A"), new AutomationChange(0, 1, "B") }, null);
            }), Request(), target);
            Assert(!result.Succeeded && target.ApplyCount == 0, "oversized result must be atomic");
        }

        private static void GenerationMismatchIsRejected()
        {
            FakeTarget target = new FakeTarget(8);
            AutomationHostResult result = Host().Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "A") }, null);
            }), Request(), target);
            Assert(!result.Succeeded && target.ApplyCount == 0 && target.Values.Count == 0, "stale request must not write");
        }

        private static void CommandExceptionIsIsolated()
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = Host().Execute(Command(delegate { throw new InvalidOperationException("boom"); }), Request(), target);
            Assert(!result.Succeeded && result.Exception is InvalidOperationException && target.ApplyCount == 0, "exception must not escape or write");
        }

        private static void BuiltInsResolveThroughRegistry()
        {
            AutomationRegistry registry = BuiltInAutomationRegistry.Create();
            IAutomationCommand command;
            Assert(registry.TryGet(ReplaceCommand.CommandId, out command) && command is ReplaceCommand,
                "replace command should be registered by stable ID");
            Assert(registry.TryGet(ReverseCommand.CommandId, out command) && command is ReverseCommand,
                "reverse command should be registered by stable ID");
            Assert(registry.TryGet(ArithmeticCommand.CommandId, out command) && command is ArithmeticCommand,
                "arithmetic command should be registered by stable ID");
            Assert(registry.TryGet(SequentialNumberCommand.CommandId, out command) && command is SequentialNumberCommand,
                "sequential command should be registered by stable ID");
            Assert(registry.TryGet(RepeatNumberCommand.CommandId, out command) && command is RepeatNumberCommand,
                "repeat command should be registered by stable ID");
        }

        private static void RegistryRejectsDuplicateIds()
        {
            AutomationRegistry registry = new AutomationRegistry();
            string error;
            Assert(registry.TryRegister(Command(delegate { return AutomationResult.Success(new AutomationChange[0], null); }), out error),
                "first registration should succeed");
            Assert(!registry.TryRegister(Command(delegate { return AutomationResult.Success(new AutomationChange[0], null); }), out error) &&
                error.IndexOf("already registered") >= 0, "duplicate ID should be rejected without replacement");
        }

        private static void RegistryKeepsDescriptorSnapshot()
        {
            AutomationRegistry registry = new AutomationRegistry();
            ChangingDescriptorCommand command = new ChangingDescriptorCommand();
            string error;
            Assert(registry.TryRegister(command, out error), "registration should read the descriptor once");
            AutomationCommandDescriptor descriptor;
            Assert(registry.TryGetDescriptor(ChangingDescriptorCommand.CommandId, out descriptor) &&
                descriptor.Id == ChangingDescriptorCommand.CommandId,
                "the registry should expose the descriptor accepted during registration");
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = new AutomationHost(10, registry).Execute(
                ChangingDescriptorCommand.CommandId, Request(), target);
            Assert(result.Succeeded && command.DescriptorReadCount == 1,
                "registered execution must not invoke extension descriptor code again");
        }

        private static void HostRejectsUnknownIds()
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = new AutomationHost(10, new AutomationRegistry()).Execute(
                "missing.command", Request(), target);
            Assert(!result.Succeeded && target.ApplyCount == 0, "unknown IDs must not execute or write");
        }

        private static void HostValidatesDescriptorParameters()
        {
            AutomationRegistry registry = new AutomationRegistry();
            ParameterCommand command = new ParameterCommand();
            string error;
            Assert(registry.TryRegister(command, out error), "parameter command registration should succeed");
            AutomationHost host = new AutomationHost(10, registry);
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult missing = host.Execute(ParameterCommand.CommandId, Request(), target);
            AutomationHostResult invalid = host.Execute(ParameterCommand.CommandId,
                Request(Parameters("count", "not-an-int", "enabled", "true")), target);
            AutomationHostResult unknown = host.Execute(ParameterCommand.CommandId,
                Request(Parameters("count", "1", "extra", "value")), target);
            Assert(!missing.Succeeded && !invalid.Succeeded && !unknown.Succeeded,
                "missing, malformed, and unknown parameters should be rejected");
            Assert(command.ExecuteCount == 0 && target.ApplyCount == 0,
                "invalid parameters must be rejected before command execution");
        }

        private static void ReplaceMatchesLeftColumnExactly()
        {
            AutomationResult result = new ReplaceCommand().Execute(CommandRequest(
                new[] { "1", "10", "1", "", "1", "1", "1", "1" },
                Parameters(ReplaceCommand.BeforeParameter, "1", ReplaceCommand.AfterParameter, "X")));
            Assert(result.Succeeded && result.Changes.Count == 2, "only exact matches in the left column should change");
            Assert(result.Changes[0].Row == 0 && result.Changes[0].Column == 0 && result.Changes[0].Value == "X", "first match expected");
            Assert(result.Changes[1].Row == 2 && result.Changes[1].Column == 0 && result.Changes[1].Value == "X", "second match expected");
        }

        private static void ReplaceRejectsEmptyValues()
        {
            AutomationResult missingBefore = new ReplaceCommand().Execute(CommandRequest(
                new[] { "1" }, Parameters(ReplaceCommand.BeforeParameter, "", ReplaceCommand.AfterParameter, "X"), 1, 1));
            AutomationResult missingAfter = new ReplaceCommand().Execute(CommandRequest(
                new[] { "1" }, Parameters(ReplaceCommand.BeforeParameter, "1", ReplaceCommand.AfterParameter, ""), 1, 1));
            Assert(!missingBefore.Succeeded && !missingAfter.Succeeded, "both empty parameters must be rejected");
        }

        private static void ReversePreservesEmptyPositions()
        {
            AutomationResult result = new ReverseCommand().Execute(CommandRequest(
                new[] { "1", "", "2", "", "3", "A", "B", "C", "D", "E" },
                new Dictionary<string, string>(), 5, 2));
            Assert(result.Succeeded && result.Changes.Count == 3, "only populated cells in the left column should be written");
            Assert(result.Changes[0].Row == 0 && result.Changes[0].Value == "3", "first value should be reversed");
            Assert(result.Changes[1].Row == 2 && result.Changes[1].Value == "2", "empty positions should remain empty");
            Assert(result.Changes[2].Row == 4 && result.Changes[2].Value == "1", "last value should be reversed");
        }

        private static void ArithmeticHandlesOperatorsAndIgnoredCells()
        {
            AutomationResult added = Arithmetic("+", "3", new[] { "1", "-2", "", "KARA" });
            AutomationResult subtracted = Arithmetic("-", "3", new[] { "1", "-2", "", "KARA" });
            AutomationResult multiplied = Arithmetic("*", "4", new[] { "2", "-3", "0", "KARA" });
            AutomationResult divided = Arithmetic("/", "2", new[] { "5", "-5", "0", "KARA" });
            Assert(Values(added) == "4,1", "addition values differ");
            Assert(Values(subtracted) == "-2,-5", "subtraction values differ");
            Assert(Values(multiplied) == "8,-12", "multiplication should skip zero and KARA");
            Assert(Values(divided) == "2,-2", "integer division should skip zero and KARA");
        }

        private static void ArithmeticRejectsInvalidCellsAtomically()
        {
            AutomationResult result = Arithmetic("+", "1", new[] { "1", "X", "2" });
            Assert(!result.Succeeded && result.Changes.Count == 0, "a nonnumeric cell must reject the entire result");
        }

        private static void ArithmeticRejectsDivisionByZero()
        {
            AutomationResult result = Arithmetic("/", "0", new[] { "1", "0", "KARA" });
            Assert(!result.Succeeded && result.Changes.Count == 0, "division by zero must be a validation failure");
        }

        private static void ArithmeticIgnoresDivisionByZeroForIgnoredCells()
        {
            AutomationResult result = Arithmetic("/", "0", new[] { "0", "", "KARA" });
            Assert(result.Succeeded && result.Changes.Count == 0, "zero, empty and KARA cells should be ignored before division");
        }

        private static void SequentialNumberPreservesBehavior()
        {
            AutomationResult compact = Sequential("10", "2", "false", 4);
            AutomationResult skipped = Sequential("10", "2", "true", 4);
            AutomationResult negative = Sequential("5", "-2", "true", 4);
            Assert(Values(compact) == "10,11", "step controls row interval when skip is off");
            Assert(Values(skipped) == "10,12", "skip should also apply step to the number");
            Assert(Values(negative) == "5,3", "negative steps should preserve legacy numbering");
        }

        private static void SequentialNumberRejectsZeroStep()
        {
            AutomationResult result = Sequential("1", "0", "false", 4);
            Assert(!result.Succeeded && result.Changes.Count == 0, "S-08 must be an input error");
        }

        private static AutomationResult Sequential(string start, string step, string skip, int rows)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add(SequentialNumberCommand.StartParameter, start);
            parameters.Add(SequentialNumberCommand.StepParameter, step);
            parameters.Add(SequentialNumberCommand.SkipParameter, skip);
            string[] values = new string[rows];
            for (int index = 0; index < values.Length; index++) values[index] = String.Empty;
            return new SequentialNumberCommand().Execute(CommandRequest(values, parameters, rows, 1));
        }

        private static void RepeatNumberHandlesParameters()
        {
            AutomationResult inserted = Repeat("1", "2", "1", "1", "0", "X", 4, 1);
            AutomationResult skipped = Repeat("1", "4", "1", "1", "1", "", 3, 1);
            AutomationResult columns = Repeat("1", "3", "1", "1", "0", "", 3, 2);
            Assert(Values(inserted) == "1,X,2,X", "insert values should alternate with numbers");
            Assert(Values(skipped) == "1,3,1", "skip behavior should match P-03");
            Assert(Values(columns) == "1,2,3,1,2,3", "numbering should continue and wrap across columns");
        }

        private static void RepeatNumberClearsAndValidates()
        {
            AutomationResult cleared = Repeat("3", "3", "1", "1", "0", "", 3, 1);
            Assert(cleared.Succeeded && Values(cleared) == "3,,", "unused selected cells should be cleared in the same result");
            Assert(!Repeat("1", "3", "0", "1", "0", "", 3, 1).Succeeded, "P-08 must reject zero interval");
            Assert(!Repeat("3", "1", "1", "1", "0", "", 3, 1).Succeeded, "P-09 must reject descending ranges");
            Assert(!Repeat("1", "3", "2", "1", "0", "", 3, 1).Succeeded, "P-10 must reject out-of-range writes atomically");
        }

        private static AutomationResult Repeat(string start, string end, string interval,
            string loop, string skip, string insert, int rows, int columns)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add(RepeatNumberCommand.StartParameter, start);
            parameters.Add(RepeatNumberCommand.EndParameter, end);
            parameters.Add(RepeatNumberCommand.RowIntervalParameter, interval);
            parameters.Add(RepeatNumberCommand.LoopParameter, loop);
            parameters.Add(RepeatNumberCommand.SkipParameter, skip);
            parameters.Add(RepeatNumberCommand.InsertParameter, insert);
            string[] values = new string[rows * columns];
            for (int index = 0; index < values.Length; index++) values[index] = "old";
            return new RepeatNumberCommand().Execute(CommandRequest(values, parameters, rows, columns));
        }

        private static void AutomationSessionReplacesOnlyItsLastResult()
        {
            SessionTarget target = new SessionTarget(7);
            AutomationSession session = new AutomationSession(Host());
            AutomationHostResult first = session.Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "first") }, null);
            }), Request(), target);
            AutomationHostResult second = session.Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "second") }, null);
            }), Request(), target);
            Assert(first.Succeeded && second.Succeeded && target.ReplaceCount == 2,
                "a session should replace its own first application");
            target.AllowPrevious = false;
            AutomationHostResult blocked = session.Execute(Command(delegate
            {
                return AutomationResult.Success(new[] { new AutomationChange(0, 0, "third") }, null);
            }), Request(), target);
            Assert(!blocked.Succeeded && target.ReplaceCount == 2,
                "an intervening operation must prevent a session from undoing arbitrary history");
        }

        private static void ExtensionDiscoveryIsolatesFailures()
        {
            string directory = CreateExtensionDirectory();
            try
            {
                AutomationExtensionLoadResult missing = new AutomationExtensionLoader().Load(
                    Path.Combine(directory, "missing"), new AutomationRegistry());
                Assert(missing.Commands.Count == 0 && missing.Diagnostics.Count == 0,
                    "a missing Extensions directory should be harmless");
                File.WriteAllText(Path.Combine(directory, "broken.dll"), "not an assembly");
                File.WriteAllText(Path.Combine(directory, "ignored.txt"), "not searched");
                string nested = Path.Combine(directory, "nested");
                Directory.CreateDirectory(nested);
                File.Copy(Assembly.GetExecutingAssembly().Location, Path.Combine(nested, "ignored.dll"));
                AutomationRegistry registry = BuiltInAutomationRegistry.Create();
                AutomationExtensionLoadResult loaded = new AutomationExtensionLoader().Load(directory, registry);
                IAutomationCommand command;
                Assert(loaded.Commands.Count == 2, "valid and runtime-fault commands should be registered");
                Assert(registry.TryGet(DiscoveryCommand.CommandId, out command), "valid external command should resolve");
                Assert(!registry.TryGet(IncompatibleDiscoveryCommand.CommandId, out command),
                    "incompatible major versions must not register");
                Assert(loaded.Diagnostics.Count >= 3, "broken DLL, constructor, and contract failures should be diagnosed");

                FakeTarget target = new FakeTarget(7);
                AutomationHostResult valid = new AutomationHost(10, registry).Execute(
                    DiscoveryCommand.CommandId, Request(), target);
                Assert(valid.Succeeded && target.ApplyCount == 1, "discovered command should execute through the host");
                AutomationHostResult faulted = new AutomationHost(10, registry).Execute(
                    RuntimeFaultDiscoveryCommand.CommandId, Request(), target);
                Assert(!faulted.Succeeded && faulted.Exception is InvalidOperationException && target.ApplyCount == 1,
                    "external execution failure must not write or escape");
            }
            finally { TryDeleteDirectory(directory); }
        }

        private static void ExtensionDiscoveryRejectsDuplicates()
        {
            string directory = CreateExtensionDirectory();
            try
            {
                AutomationRegistry registry = BuiltInAutomationRegistry.Create();
                AutomationExtensionLoader loader = new AutomationExtensionLoader();
                AutomationExtensionLoadResult first = loader.Load(directory, registry);
                AutomationExtensionLoadResult second = loader.Load(directory, registry);
                Assert(first.Commands.Count == 2 && second.Commands.Count == 0,
                    "a repeated discovery must not replace registered IDs");
                Assert(second.Diagnostics.Count >= 2, "duplicate IDs should be diagnosed");
            }
            finally { TryDeleteDirectory(directory); }
        }

        private static string CreateExtensionDirectory()
        {
            string directory = Path.Combine(Path.GetTempPath(), "aeiou-extension-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            File.Copy(Assembly.GetExecutingAssembly().Location, Path.Combine(directory, "fixtures.dll"));
            return directory;
        }

        private static void TryDeleteDirectory(string directory)
        {
            try { Directory.Delete(directory, true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static AutomationResult Arithmetic(string operation, string operand, string[] values)
        {
            return new ArithmeticCommand().Execute(CommandRequest(values,
                Parameters(ArithmeticCommand.OperatorParameter, operation, ArithmeticCommand.OperandParameter, operand),
                values.Length, 1));
        }

        private static string Values(AutomationResult result)
        {
            List<string> values = new List<string>();
            foreach (AutomationChange change in result.Changes) values.Add(change.Value);
            return String.Join(",", values.ToArray());
        }

        private static Dictionary<string, string> Parameters(string firstKey, string firstValue,
            string secondKey, string secondValue)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add(firstKey, firstValue);
            parameters.Add(secondKey, secondValue);
            return parameters;
        }

        private static AutomationRequest CommandRequest(string[] values, IDictionary<string, string> parameters)
        {
            return CommandRequest(values, parameters, 4, 2);
        }

        private static AutomationRequest CommandRequest(string[] values, IDictionary<string, string> parameters,
            int rows, int columns)
        {
            List<AutomationCell> cells = new List<AutomationCell>();
            int index = 0;
            for (int column = 0; column < columns; column++)
                for (int row = 0; row < rows; row++)
                    cells.Add(new AutomationCell(row, column, values[index++]));
            return new AutomationRequest(rows, columns, 1, new AutomationSelection(0, 0, rows, columns),
                cells, parameters, "KARA");
        }

        private static void AssertRejectedWithoutWrites(IEnumerable<AutomationChange> changes)
        {
            FakeTarget target = new FakeTarget(7);
            AutomationHostResult result = Host().Execute(Command(delegate { return AutomationResult.Success(changes, null); }), Request(), target);
            Assert(!result.Succeeded && target.ApplyCount == 0 && target.Values.Count == 0, "invalid set must write zero cells");
        }

        private static AutomationHost Host() { return new AutomationHost(10); }

        private static AutomationRequest Request()
        {
            return Request(new Dictionary<string, string>());
        }

        private static AutomationRequest Request(IDictionary<string, string> parameters)
        {
            return new AutomationRequest(2, 2, 7, new AutomationSelection(0, 0, 2, 2),
                new[] { new AutomationCell(0, 0, "") }, parameters, "KARA");
        }

        private static IAutomationCommand Command(Func<AutomationResult> execute) { return new FakeCommand(execute); }

        private static void Run(string name, Action test)
        {
            try { test(); Console.WriteLine("PASS: " + name); }
            catch (Exception exception) { failures++; Console.Error.WriteLine("FAIL: " + name + " - " + exception.Message); }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private sealed class FakeCommand : IAutomationCommand
        {
            private readonly Func<AutomationResult> execute;
            public FakeCommand(Func<AutomationResult> execute) { this.execute = execute; }
            public AutomationCommandDescriptor Descriptor
            {
                get { return new AutomationCommandDescriptor("test.command", "Test", 1, 0, new AutomationParameterDefinition[0]); }
            }
            public AutomationResult Execute(AutomationRequest request) { return execute(); }
        }

        private sealed class FakeClipboardWriter : IAeClipboardWriter
        {
            public int WriteCount { get; private set; }
            public string Text { get; private set; }

            public void SetText(string text)
            {
                WriteCount++;
                Text = text;
            }
        }

        private sealed class ParameterCommand : IAutomationCommand
        {
            public const string CommandId = "test.parameters";
            public int ExecuteCount;
            public AutomationCommandDescriptor Descriptor
            {
                get
                {
                    return new AutomationCommandDescriptor(CommandId, "Parameters", 1, 0, new[]
                    {
                        new AutomationParameterDefinition("count", "Count", AutomationParameterType.Int32,
                            "1", true, 0, 10, null),
                        new AutomationParameterDefinition("enabled", "Enabled", AutomationParameterType.Boolean,
                            "false", true, null, null, null)
                    });
                }
            }
            public AutomationResult Execute(AutomationRequest request)
            {
                ExecuteCount++;
                return AutomationResult.Success(new AutomationChange[0], null);
            }
        }

        private sealed class ChangingDescriptorCommand : IAutomationCommand
        {
            public const string CommandId = "test.changing-descriptor";
            public int DescriptorReadCount;
            public AutomationCommandDescriptor Descriptor
            {
                get
                {
                    DescriptorReadCount++;
                    if (DescriptorReadCount > 1) throw new InvalidOperationException("descriptor read twice");
                    return new AutomationCommandDescriptor(CommandId, "Changing descriptor", 1, 0,
                        new AutomationParameterDefinition[0]);
                }
            }
            public AutomationResult Execute(AutomationRequest request)
            {
                return AutomationResult.Success(new AutomationChange[0], null);
            }
        }

        private sealed class FakeTarget : IAutomationChangeTarget
        {
            private readonly long generation;
            public readonly Dictionary<string, string> Values = new Dictionary<string, string>();
            public int ApplyCount;
            public FakeTarget(long generation) { this.generation = generation; }
            public bool TryApply(long expectedGeneration, IList<AutomationChange> changes, string operationName)
            {
                if (expectedGeneration != generation) return false;
                ApplyCount++;
                foreach (AutomationChange change in changes) Values[change.Row + ":" + change.Column] = change.Value;
                return true;
            }
        }

        private sealed class SessionTarget : IAutomationSessionTarget
        {
            private readonly long generation;
            private object lastApplication;
            public bool AllowPrevious = true;
            public int ReplaceCount;

            public SessionTarget(long generation) { this.generation = generation; }

            public bool TryReplace(long expectedGeneration, object previousApplication,
                IList<AutomationChange> changes, string operationName, out object application)
            {
                application = null;
                if (expectedGeneration != generation || previousApplication != null &&
                    (!AllowPrevious || !Object.ReferenceEquals(previousApplication, lastApplication))) return false;
                application = new object();
                lastApplication = application;
                ReplaceCount++;
                return true;
            }
        }
    }

    public sealed class DiscoveryCommand : IAutomationCommand
    {
        public const string CommandId = "test.external.valid";
        public AutomationCommandDescriptor Descriptor { get { return DescriptorFor(CommandId, 1); } }
        public AutomationResult Execute(AutomationRequest request)
        {
            return AutomationResult.Success(new[] { new AutomationChange(0, 0, "external") }, null);
        }
        internal static AutomationCommandDescriptor DescriptorFor(string id, int major)
        {
            return new AutomationCommandDescriptor(id, "External fixture", major, 0,
                new AutomationParameterDefinition[0]);
        }
    }

    public sealed class IncompatibleDiscoveryCommand : IAutomationCommand
    {
        public const string CommandId = "test.external.incompatible";
        public AutomationCommandDescriptor Descriptor { get { return DiscoveryCommand.DescriptorFor(CommandId, 2); } }
        public AutomationResult Execute(AutomationRequest request) { return AutomationResult.Success(new AutomationChange[0], null); }
    }

    public sealed class ConstructorFaultDiscoveryCommand : IAutomationCommand
    {
        public ConstructorFaultDiscoveryCommand() { throw new InvalidOperationException("constructor fixture"); }
        public AutomationCommandDescriptor Descriptor { get { return DiscoveryCommand.DescriptorFor("test.external.constructor", 1); } }
        public AutomationResult Execute(AutomationRequest request) { return AutomationResult.Success(new AutomationChange[0], null); }
    }

    public sealed class RuntimeFaultDiscoveryCommand : IAutomationCommand
    {
        public const string CommandId = "test.external.runtime";
        public AutomationCommandDescriptor Descriptor { get { return DiscoveryCommand.DescriptorFor(CommandId, 1); } }
        public AutomationResult Execute(AutomationRequest request) { throw new InvalidOperationException("runtime fixture"); }
    }

    public abstract class AbstractDiscoveryCommand : IAutomationCommand
    {
        public abstract AutomationCommandDescriptor Descriptor { get; }
        public abstract AutomationResult Execute(AutomationRequest request);
    }
}
