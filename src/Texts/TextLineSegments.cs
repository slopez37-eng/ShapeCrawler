using System;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents text split into line segments.
/// </summary>
/// <param name="text">The text to split.</param>
internal sealed class TextLineSegments(string text)
{
    /// <summary>
    ///     Converts text to line segments.
    /// </summary>
    internal string[] ToArray()
    {
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
    }
}