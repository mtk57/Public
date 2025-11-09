using System;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpProgressInfo
    {
        public TagJumpProgressInfo(string phase, int processedCount, int totalCount,
            int processedFileCount = 0, int totalFileCount = 0)
        {
            Phase = phase ?? string.Empty;
            ProcessedCount = processedCount < 0 ? 0 : processedCount;
            TotalCount = totalCount < 0 ? 0 : totalCount;
            ProcessedFileCount = processedFileCount < 0 ? 0 : processedFileCount;
            TotalFileCount = totalFileCount < 0 ? 0 : totalFileCount;
        }

        public string Phase { get; }

        public int ProcessedCount { get; }

        public int TotalCount { get; }

        public int ProcessedFileCount { get; }

        public int TotalFileCount { get; }

        public bool IsDeterminate => TotalCount > 0;

        public bool HasFileProgress => TotalFileCount > 0;

        public int CalculatePercentage()
        {
            if (!IsDeterminate || TotalCount == 0)
            {
                return 0;
            }

            var percentage = (double)ProcessedCount / TotalCount * 100;
            if (percentage < 0)
            {
                percentage = 0;
            }
            else if (percentage > 100)
            {
                percentage = 100;
            }

            return (int)Math.Round(percentage);
        }
    }
}
