using System;
using System.Collections.Generic;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpEmbeddingResult
    {
        public TagJumpEmbeddingResult(int updatedFileCount, int updatedCallCount,
            IReadOnlyList<string> updatedFilePaths, int failureCount,
            IReadOnlyList<TagJumpFailureDetail> failureDetails)
        {
            UpdatedFileCount = updatedFileCount;
            UpdatedCallCount = updatedCallCount;
            UpdatedFilePaths = updatedFilePaths ?? Array.Empty<string>();
            FailureCount = failureCount < 0 ? 0 : failureCount;
            FailureDetails = failureDetails ?? Array.Empty<TagJumpFailureDetail>();
        }

        public int UpdatedFileCount { get; }

        public int UpdatedCallCount { get; }

        public IReadOnlyList<string> UpdatedFilePaths { get; }

        public int FailureCount { get; }

        public IReadOnlyList<TagJumpFailureDetail> FailureDetails { get; }
    }
}
