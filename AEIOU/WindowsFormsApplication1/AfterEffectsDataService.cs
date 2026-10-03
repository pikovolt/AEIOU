using System;
using System.Collections.Generic;
using System.Text;

namespace AEIOU
{
    internal sealed class AfterEffectsKeyframe
    {
        public AfterEffectsKeyframe(int frame, double value)
        {
            Frame = frame;
            Value = value;
        }

        public int Frame { get; private set; }
        public double Value { get; private set; }
    }

    internal sealed class AfterEffectsPasteData
    {
        public AfterEffectsPasteData(int fps, List<AfterEffectsKeyframe> keyframes)
        {
            Fps = fps;
            Keyframes = keyframes;
        }

        public int Fps { get; private set; }
        public List<AfterEffectsKeyframe> Keyframes { get; private set; }
    }

    /// <summary>After Effects と交換する文字列の生成・解析を担当する。</summary>
    internal sealed class AfterEffectsDataService
    {
        public string CreateKeyframeData(string version, int fps, int firstFrame, bool isDirect,
            int rowCount, Func<int, string> getCellValue, Func<int, bool> isExcludedRow)
        {
            StringBuilder text = new StringBuilder();
            text.Append("Adobe After Effects ").Append(version).Append(" Keyframe Data\r\n\r\n");
            text.Append("\tUnits Per Second\t").Append(fps).Append("\r\n");
            text.Append("\tSource Width\t640\r\n\tSource Height\t480\r\n");
            text.Append("\tSource Pixel Aspect Ratio\t1\r\n\tComp Pixel Aspect Ratio\t1\r\n\r\n");
            text.Append("Time Remap\r\n\tFrame\tseconds\r\n");

            int excludedCount = 0;
            for (int row = 0; row < rowCount; row++)
            {
                if (isExcludedRow(row))
                {
                    excludedCount++;
                    continue;
                }

                string value = getCellValue(row);
                if (value.Length == 0)
                {
                    continue;
                }

                double timing = Int32.Parse(value);
                if (!isDirect)
                {
                    timing = (timing - firstFrame) / fps;
                }
                text.Append('\t').Append(row - excludedCount).Append('\t').Append(timing.ToString("g6")).Append("\r\n");
            }

            text.Append("\r\n\r\nEnd of Keyframe Data\r\n");
            return text.ToString();
        }

        public string CreateScriptData(int fps, int firstFrame, bool isDirect, int rowCount,
            Func<int, string> getCellValue, Func<int, bool> isExcludedRow)
        {
            StringBuilder text = new StringBuilder();
            text.Append("{property:'Time Remap',scale:").Append(fps.ToString("f")).Append(",keys:[");
            int excludedCount = 0;
            for (int row = 0; row < rowCount; row++)
            {
                if (isExcludedRow(row))
                {
                    excludedCount++;
                    continue;
                }

                string value = getCellValue(row);
                if (value.Length == 0)
                {
                    continue;
                }

                double timing = Int32.Parse(value);
                if (!isDirect)
                {
                    timing -= firstFrame;
                }
                text.Append("{'t':").Append(row - excludedCount).Append(",'v':[").Append(timing).Append("]},");
            }
            text.Append("]}");
            return text.ToString();
        }

        public bool TryParseKeyframeData(string clipboardText, out AfterEffectsPasteData data, out string errorMessage)
        {
            data = null;
            errorMessage = null;
            string[] lines = clipboardText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            if (lines.Length == 0 || lines[0].IndexOf("Adobe After Effects ") < 0 ||
                lines[0].IndexOf("Keyframe Data") < 0)
            {
                errorMessage = "ヘッダ : 未対応のヘッダ.";
                return false;
            }

            if (lines.Length <= 2)
            {
                errorMessage = "Units Per Second : 目的の情報が見つからない.";
                return false;
            }
            string[] fpsFields = lines[2].Split('\t');
            int fps;
            if (lines[2].IndexOf("\tUnits Per Second") < 0 || fpsFields.Length < 3 ||
                !Int32.TryParse(fpsFields[2], out fps))
            {
                errorMessage = "Units Per Second : 目的の情報が見つからない.";
                return false;
            }

            int index = 3;
            while (index < lines.Length && lines[index].IndexOf("Time Remap") < 0)
            {
                index++;
            }
            if (index >= lines.Length)
            {
                errorMessage = "Time Remap : 目的の情報が見つからない.";
                return false;
            }

            List<AfterEffectsKeyframe> keyframes = new List<AfterEffectsKeyframe>();
            for (index += 2; index < lines.Length; index++)
            {
                string[] fields = lines[index].Split('\t');
                if (fields.Length <= 1)
                {
                    break;
                }
                int frame;
                double value;
                if (fields.Length < 3 || !Int32.TryParse(fields[1], out frame) || !Double.TryParse(fields[2], out value))
                {
                    errorMessage = "Time Remap : キーフレーム情報を読み取れませんでした.";
                    return false;
                }
                keyframes.Add(new AfterEffectsKeyframe(frame, value));
            }

            data = new AfterEffectsPasteData(fps, keyframes);
            return true;
        }
    }
}
