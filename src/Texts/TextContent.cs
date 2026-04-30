using System;
using System.Linq;
using ShapeCrawler.Paragraphs;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents a plain text content.
/// </summary>
internal sealed class TextContent(
    string text,
    IParagraphCollection paragraphs,
    Func<AutofitType> getAutofitType,
    Action<string> shrinkFont,
    Action applyResize)
{
    /// <summary>
    ///     Applies the text content to the paragraphs.
    /// </summary>
    internal void ApplyTo()
    {
        var paragraphsList = paragraphs.ToArray();
        var firstParagraph = paragraphsList.FirstOrDefault();

        var latinNameToPreserve = GetLatinNameToPreserve(firstParagraph);
        var colorHexToPreserve = GetFontColorHexToPreserve(firstParagraph);

        // Clear existing content and ensure we have a first paragraph
        firstParagraph = this.PrepareContainer(firstParagraph, paragraphsList);

        var paragraphLines = text.Split([Environment.NewLine], StringSplitOptions.None);
        this.AddToParagraphs(paragraphLines, firstParagraph, latinNameToPreserve);
        this.ApplyFontColorIfNeeded(paragraphLines, colorHexToPreserve);

        this.ApplyFormatting();
    }

    private static string? GetLatinNameToPreserve(IParagraph? firstParagraph)
    {
        var aRunProperties = FirstRunPropertiesOrNull(firstParagraph);
        return aRunProperties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
    }

    private static string? GetFontColorHexToPreserve(IParagraph? firstParagraph)
    {
        var aRunProperties = FirstRunPropertiesOrNull(firstParagraph);
        if (aRunProperties?.GetFirstChild<A.SolidFill>() == null)
        {
            return null;
        }

        var firstPortion = firstParagraph?.Portions.FirstOrDefault();
        return firstPortion?.Font?.Color.Hex;
    }

    // Read Open XML directly so we only see properties explicitly set on the run.
    // The IFont API resolves through layout/master/theme inheritance, which would
    // cause inherited values to be re-stamped on new runs and break theme inheritance.
    private static A.RunProperties? FirstRunPropertiesOrNull(IParagraph? firstParagraph)
    {
        return firstParagraph?.Portions.FirstOrDefault() is TextParagraphPortion textParagraphPortion
            ? textParagraphPortion.AText.Parent?.GetFirstChild<A.RunProperties>()
            : null;
    }

    private static void ApplyLatinNameIfNeeded(IParagraphPortion portion, string? latinNameToPreserve)
    {
        if (latinNameToPreserve != null && portion.Font != null)
        {
            portion.Font.LatinName = latinNameToPreserve;
        }
    }

    private void ApplyFontColorIfNeeded(string[] paragraphLines, string? colorHexToPreserve)
    {
        if (colorHexToPreserve == null)
        {
            return;
        }

        for (int i = 0; i < paragraphLines.Length; i++)
        {
            var portion = paragraphs[i].Portions.Last();
            portion.Font!.Color.Set(colorHexToPreserve);
        }
    }

    private IParagraph PrepareContainer(IParagraph? firstParagraph, IParagraph[] paragraphsList)
    {
        if (firstParagraph == null)
        {
            paragraphs.Add();
            return paragraphs.First();
        }

        foreach (var paragraph in paragraphsList.Skip(1))
        {
            paragraph.Remove();
        }

        foreach (var portion in firstParagraph.Portions.ToList())
        {
            portion.Remove();
        }

        return firstParagraph;
    }

    private void AddToParagraphs(string[] paragraphLines, IParagraph firstParagraph, string? latinNameToPreserve)
    {
        if (paragraphLines.Length <= 0)
        {
            return;
        }

        firstParagraph.Portions.AddText(paragraphLines[0]);
        ApplyLatinNameIfNeeded(firstParagraph.Portions.Last(), latinNameToPreserve);

        for (int i = 1; i < paragraphLines.Length; i++)
        {
            paragraphs.Add();
            paragraphs[i].Portions.AddText(paragraphLines[i]);
            ApplyLatinNameIfNeeded(paragraphs[i].Portions.Last(), latinNameToPreserve);
        }
    }

    private void ApplyFormatting()
    {
        if (getAutofitType() == AutofitType.Shrink)
        {
            shrinkFont(text);
        }

        applyResize();
    }
}