using System;

namespace SimpleSqlAdjuster
{
    internal class SqlProcessingException : Exception
    {
        public SqlProcessingException(string message, int line, int column)
            : base(message)
        {
            Line = line;
            Column = column;
        }

        public SqlProcessingException(string message, int line, int column, Exception innerException)
            : base(message, innerException)
        {
            Line = line;
            Column = column;
        }

        public int Line { get; }

        public int Column { get; }

        public string ToDisplayMessage()
        {
            return string.Format("行 {0}, 列 {1}: {2}", Line, Column, Message);
        }
    }
}
