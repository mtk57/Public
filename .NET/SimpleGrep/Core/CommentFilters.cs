using System;
using System.IO;
using System.Text;

namespace SimpleGrep.Core
{
    internal interface ICommentFilter
    {
        string RemoveComments(string line);
    }

    internal static class CommentFilterFactory
    {
        public static ICommentFilter Create(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            if (string.Equals(extension, ".java", StringComparison.OrdinalIgnoreCase))
            {
                return new JavaCommentFilter();
            }

            return null;
        }
    }

    internal sealed class JavaCommentFilter : ICommentFilter
    {
        private bool inBlockComment;

        public string RemoveComments(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            bool inString = false;
            bool inChar = false;
            bool escape = false;

            for (int i = 0; i < line.Length; i++)
            {
                char current = line[i];
                char next = i + 1 < line.Length ? line[i + 1] : '\0';

                if (inBlockComment)
                {
                    if (current == '*' && next == '/')
                    {
                        inBlockComment = false;
                        i++;
                    }
                    continue;
                }

                if (inString)
                {
                    sb.Append(current);
                    if (escape)
                    {
                        escape = false;
                    }
                    else
                    {
                        if (current == '\\')
                        {
                            escape = true;
                        }
                        else if (current == '"')
                        {
                            inString = false;
                        }
                    }
                    continue;
                }

                if (inChar)
                {
                    sb.Append(current);
                    if (escape)
                    {
                        escape = false;
                    }
                    else
                    {
                        if (current == '\\')
                        {
                            escape = true;
                        }
                        else if (current == '\'')
                        {
                            inChar = false;
                        }
                    }
                    continue;
                }

                if (current == '/' && next == '/')
                {
                    break;
                }

                if (current == '/' && next == '*')
                {
                    inBlockComment = true;
                    i++;
                    continue;
                }

                if (current == '"')
                {
                    inString = true;
                    sb.Append(current);
                    escape = false;
                    continue;
                }

                if (current == '\'')
                {
                    inChar = true;
                    sb.Append(current);
                    escape = false;
                    continue;
                }

                sb.Append(current);
            }

            return sb.ToString();
        }
    }
}
