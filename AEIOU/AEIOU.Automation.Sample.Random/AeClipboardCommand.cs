using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using AEIOU.Automation;

namespace AEIOU.Automation.Sample.Random
{
    /// <summary>Example external command that writes AE keyframe data directly to the clipboard.</summary>
    public sealed class AeClipboardCommand : IAutomationCommand
    {
        public const string CommandId = "sample.ae-clipboard";
        public const string VersionParameter = "version";
        public const string FpsParameter = "fps";
        public const string FirstFrameParameter = "first_frame";
        public const string DirectParameter = "direct";

        private readonly IAeClipboardWriter clipboard;

        public AeClipboardCommand() : this(new WindowsClipboardWriter()) { }

        public AeClipboardCommand(IAeClipboardWriter clipboard)
        {
            if (clipboard == null) throw new ArgumentNullException("clipboard");
            this.clipboard = clipboard;
        }

        public AutomationCommandDescriptor Descriptor
        {
            get
            {
                return new AutomationCommandDescriptor(CommandId, "AEコピー（サンプル）",
                    AutomationContract.MajorVersion, AutomationContract.MinorVersion, new[]
                    {
                        new AutomationParameterDefinition(VersionParameter, "AEバージョン",
                            AutomationParameterType.String, "9.0", true, null, null, null),
                        new AutomationParameterDefinition(FpsParameter, "FPS",
                            AutomationParameterType.Int32, "24", true, 1, Int32.MaxValue, null),
                        new AutomationParameterDefinition(FirstFrameParameter, "先頭フレーム",
                            AutomationParameterType.Int32, "1", true, Int32.MinValue, Int32.MaxValue, null),
                        new AutomationParameterDefinition(DirectParameter, "値を秒へ変換しない",
                            AutomationParameterType.Boolean, "false", true, null, null, null)
                    });
            }
        }

        public AutomationResult Execute(AutomationRequest request)
        {
            if (request == null) throw new ArgumentNullException("request");
            if (request.Selection.ColumnCount != 1)
                return AutomationResult.Failure("AEコピーでは1列だけを選択してください。");

            string version;
            string parameterValue;
            int fps;
            int firstFrame;
            bool direct;
            if (!request.Parameters.TryGetValue(VersionParameter, out version) ||
                String.IsNullOrEmpty(version) ||
                !request.Parameters.TryGetValue(FpsParameter, out parameterValue) ||
                !Int32.TryParse(parameterValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out fps) ||
                !request.Parameters.TryGetValue(FirstFrameParameter, out parameterValue) ||
                !Int32.TryParse(parameterValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out firstFrame) ||
                !request.Parameters.TryGetValue(DirectParameter, out parameterValue) ||
                !Boolean.TryParse(parameterValue, out direct))
                return AutomationResult.Failure("AEバージョン、FPS、先頭フレームの入力を確認してください。");
            if (fps <= 0) return AutomationResult.Failure("FPSには1以上を指定してください。");

            string text;
            string error;
            if (!TryCreateClipboardText(request, version, fps, firstFrame, direct, out text, out error))
                return AutomationResult.Failure(error);

            clipboard.SetText(text);
            return AutomationResult.Success(new AutomationChange[0],
                "AEキーフレームデータをクリップボードへコピーしました。");
        }

        private static bool TryCreateClipboardText(AutomationRequest request, string version, int fps,
            int firstFrame, bool direct, out string text, out string error)
        {
            Dictionary<int, string> values = new Dictionary<int, string>();
            foreach (AutomationCell cell in request.Cells) values[cell.Row] = cell.Value;

            StringBuilder builder = new StringBuilder();
            builder.Append("Adobe After Effects ").Append(version).Append(" Keyframe Data\r\n\r\n");
            builder.Append("\tUnits Per Second\t").Append(fps.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
            builder.Append("\tSource Width\t640\r\n\tSource Height\t480\r\n");
            builder.Append("\tSource Pixel Aspect Ratio\t1\r\n\tComp Pixel Aspect Ratio\t1\r\n\r\n");
            builder.Append("Time Remap\r\n\tFrame\tseconds\r\n");

            for (int offset = 0; offset < request.Selection.RowCount; offset++)
            {
                string value;
                values.TryGetValue(request.Selection.Top + offset, out value);
                if (String.IsNullOrEmpty(value) || value == request.EmptyCellValue) continue;
                int frameValue;
                if (!Int32.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out frameValue))
                {
                    text = null;
                    error = "セルの値を数値に変換できませんでした。";
                    return false;
                }
                double output = direct ? frameValue : (frameValue - firstFrame) / (double)fps;
                builder.Append('\t').Append(offset).Append('\t')
                    .Append(output.ToString("g6", CultureInfo.InvariantCulture)).Append("\r\n");
            }
            builder.Append("\r\n\r\nEnd of Keyframe Data\r\n");
            text = builder.ToString();
            error = null;
            return true;
        }

        private sealed class WindowsClipboardWriter : IAeClipboardWriter
        {
            public void SetText(string text)
            {
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    try
                    {
                        Clipboard.Clear();
                        Clipboard.SetText(text);
                        return;
                    }
                    catch (ExternalException)
                    {
                        if (attempt == 4) throw;
                        Thread.Sleep(100);
                    }
                }
            }
        }
    }

    public interface IAeClipboardWriter
    {
        void SetText(string text);
    }
}
