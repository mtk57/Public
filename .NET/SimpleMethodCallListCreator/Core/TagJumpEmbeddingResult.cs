using System;
using System.Collections.Generic;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpEmbeddingResult
    {
        public TagJumpEmbeddingResult(int updatedFileCount, int updatedCallCount,
            IReadOnlyList<string> updatedFilePaths)
        {
            UpdatedFileCount = updatedFileCount;
            UpdatedCallCount = updatedCallCount;
            UpdatedFilePaths = updatedFilePaths ?? Array.Empty<string>();
        }

        public int UpdatedFileCount { get; }

        public int UpdatedCallCount { get; }

        public IReadOnlyList<string> UpdatedFilePaths { get; }
    }
}
