using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace AEIOU
{
    internal sealed class StsCellChange
    {
        public StsCellChange(int column, int row, string value)
        {
            Column = column;
            Row = row;
            Value = value;
        }

        public int Column { get; private set; }
        public int Row { get; private set; }
        public string Value { get; private set; }
    }

    internal sealed class StsDocument
    {
        public StsDocument(int columnCount, int rowCount, List<StsCellChange> cellChanges, List<string> headers)
        {
            ColumnCount = columnCount;
            RowCount = rowCount;
            CellChanges = cellChanges;
            Headers = headers;
        }

        public int ColumnCount { get; private set; }
        public int RowCount { get; private set; }
        public List<StsCellChange> CellChanges { get; private set; }
        public List<string> Headers { get; private set; }
    }

    /// <summary>STS のバイナリ形式だけを担当し、画面やシートモデルには依存しない。</summary>
    internal sealed class StsFileService
    {
        private static readonly byte[] Header =
        {
            0x11, (byte)'S', (byte)'h', (byte)'i', (byte)'r', (byte)'a', (byte)'h', (byte)'e',
            (byte)'i', (byte)'T', (byte)'i', (byte)'m', (byte)'e', (byte)'S', (byte)'h', (byte)'e', (byte)'e', (byte)'t'
        };

        private static readonly Encoding StsEncoding = Encoding.GetEncoding("Shift_JIS");

        public void Save(string path, int columnCount, int rowCount,
            Func<int, int, string> getCellValue, Func<int, string> getHeaderValue)
        {
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                stream.Write(Header, 0, Header.Length);
                stream.WriteByte((byte)columnCount);
                WriteUInt32(stream, (UInt32)rowCount);

                for (int column = 0; column < columnCount; column++)
                {
                    UInt16 current = 0;
                    for (int row = 0; row < rowCount; row++)
                    {
                        string value = getCellValue(column, row);
                        if (value.Length > 0)
                        {
                            UInt16 parsed = 0;
                            UInt16.TryParse(value, out parsed);
                            current = parsed;
                        }
                        WriteUInt16(stream, current);
                    }
                }

                for (int column = 0; column < columnCount; column++)
                {
                    byte[] name = StsEncoding.GetBytes(getHeaderValue(column));
                    stream.WriteByte((byte)name.Length);
                    stream.Write(name, 0, name.Length);
                }
            }
        }

        public StsDocument Load(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                byte[] actualHeader = ReadBytes(stream, Header.Length);
                for (int i = 0; i < Header.Length; i++)
                {
                    if (actualHeader[i] != Header[i])
                    {
                        throw new InvalidDataException("Unsupported STS header.");
                    }
                }

                int columnCount = stream.ReadByte();
                if (columnCount < 0)
                {
                    throw new EndOfStreamException();
                }
                UInt32 rawRowCount = BitConverter.ToUInt32(ReadBytes(stream, sizeof(UInt32)), 0);
                if (rawRowCount > Int32.MaxValue)
                {
                    throw new InvalidDataException("STS row count is too large.");
                }
                int rowCount = (int)rawRowCount;

                List<StsCellChange> changes = new List<StsCellChange>();
                for (int column = 0; column < columnCount; column++)
                {
                    UInt16 current = 0;
                    for (int row = 0; row < rowCount; row++)
                    {
                        UInt16 value = BitConverter.ToUInt16(ReadBytes(stream, sizeof(UInt16)), 0);
                        if (current != value)
                        {
                            current = value;
                            changes.Add(new StsCellChange(column, row, value.ToString()));
                        }
                    }
                }

                List<string> headers = new List<string>();
                for (int column = 0; column < columnCount; column++)
                {
                    int nameLength = stream.ReadByte();
                    if (nameLength < 0)
                    {
                        throw new EndOfStreamException();
                    }
                    headers.Add(StsEncoding.GetString(ReadBytes(stream, nameLength)));
                }

                return new StsDocument(columnCount, rowCount, changes, headers);
            }
        }

        private static byte[] ReadBytes(Stream stream, int count)
        {
            byte[] result = new byte[count];
            int offset = 0;
            while (offset < count)
            {
                int read = stream.Read(result, offset, count - offset);
                if (read == 0)
                {
                    throw new EndOfStreamException();
                }
                offset += read;
            }
            return result;
        }

        private static void WriteUInt16(Stream stream, UInt16 value)
        {
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt32(Stream stream, UInt32 value)
        {
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }
    }
}
