using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents collection of paragraph text portions.
/// </summary>
public interface IParagraphPortions : IEnumerable<IParagraphPortion>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    IParagraphPortion this[int index] { get; }

    /// <summary>
    ///     Adds text portion.
    /// </summary>
    void AddText(string text);

    /// <summary>
    ///     Adds Line Break.
    /// </summary>
    void AddLineBreak();
}

internal sealed class ParagraphPortions(A.Paragraph aParagraph) : IParagraphPortions
{
    public int Count => this.GetPortions().Count;

    public IParagraphPortion this[int index] => this.GetPortions()[index];

    public void AddText(string text)
    {
        var lastRunOrBreak = aParagraph.LastOrDefault(p => p is A.Run or A.Break);
        var textPortions = this.GetPortions().OfType<TextParagraphPortion>();
        var aTextParent = textPortions.LastOrDefault()?.AText.Parent;
        if (aTextParent is null)
        {
            var aRunProperties = new A.RunProperties { Language = "en-US", Dirty = false };

            // Skip the historical 14pt stamp only for placeholders that actually
            // inherit from a matching layout placeholder. Orphan <p:ph> tags with
            // no layout counterpart fall back to the default, matching upstream.
            if (!InheritsFromLayoutPlaceholder(aParagraph))
            {
                aRunProperties.FontSize = 1400;
            }

            var aText = new A.Text { Text = string.Empty };
            aTextParent = new A.Run(aRunProperties, aText);
        }

        this.AddText(ref lastRunOrBreak, aTextParent, text);
    }

    public void AddLineBreak()
    {
        var lastARunOrABreak = aParagraph.Last();
        lastARunOrABreak.InsertAfterSelf(new A.Break());
    }

    public IEnumerator<IParagraphPortion> GetEnumerator() => this.GetPortions().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private static bool InheritsFromLayoutPlaceholder(A.Paragraph aParagraph)
    {
        var pShape = aParagraph.Ancestors<P.Shape>().FirstOrDefault();
        var slidePh = pShape?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return false;
        }

        var partRoot = aParagraph.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
        if (partRoot?.OpenXmlPart is not SlidePart slidePart)
        {
            return false;
        }

        var layoutShapeTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
        if (layoutShapeTree == null)
        {
            return false;
        }

        return new SCPShapeTree(layoutShapeTree).ReferencedPShapeOrNull(slidePh) != null;
    }

    private void AddText(ref OpenXmlElement? lastElement, OpenXmlElement aTextParent, string text)
    {
        var newARun = (A.Run)aTextParent.CloneNode(true);
        newARun.Text!.Text = text;
        if (lastElement == null)
        {
            var apPr = aParagraph.GetFirstChild<A.ParagraphProperties>();
            lastElement = apPr != null ? apPr.InsertAfterSelf(newARun) : aParagraph.InsertAt(newARun, 0);
        }
        else
        {
            lastElement = lastElement.InsertAfterSelf(newARun);
        }
    }

    private List<IParagraphPortion> GetPortions()
    {
        var portions = new List<IParagraphPortion>();
        foreach (var aParagraphElement in aParagraph.Elements())
        {
            switch (aParagraphElement)
            {
                case A.Run aRun:
                    var runPortion = new TextParagraphPortion(aRun);
                    portions.Add(runPortion);
                    break;
                case A.Field aField:
                    var fieldPortion = new Field(aField);
                    portions.Add(fieldPortion);
                    break;
                case A.Break aBreak:
                    var lineBreak = new ParagraphLineBreak(aBreak);
                    portions.Add(lineBreak);
                    break;
                default:
                    continue;
            }
        }

        return portions;
    }
}