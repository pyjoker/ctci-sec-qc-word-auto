using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace WordAutoTool.Services;

public class WordProcessingService
{
    /// <summary>
    /// Replace text box dates and inline body dates. Images are handled separately via WordComService.
    /// </summary>
    public byte[] Process(byte[] docxBytes, string rocDate)
    {
        using var ms = new MemoryStream();
        ms.Write(docxBytes);
        ms.Position = 0;

        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            ReplaceTextBoxDates(doc, rocDate);
            ReplaceInlineBodyDates(body, rocDate);
        }

        return ms.ToArray();
    }

    // ─── 1. Replace ALL text box content with rocDate ─────────────────────────

    private static void ReplaceTextBoxDates(WordprocessingDocument doc, string rocDate)
    {
        foreach (var txbxContent in doc.MainDocumentPart!.Document.Body!
            .Descendants<W.TextBoxContent>()
            .ToList())
        {
            SetSingleRunText(txbxContent, rocDate);
        }
    }

    private static void SetSingleRunText(OpenXmlElement container, string rocDate)
    {
        var paragraphs = container.Elements<W.Paragraph>().ToList();
        if (paragraphs.Count == 0) return;

        var firstPara = paragraphs[0];

        W.RunProperties? rpr = firstPara
            .Descendants<W.Run>()
            .FirstOrDefault()
            ?.GetFirstChild<W.RunProperties>()
            ?.CloneNode(true) as W.RunProperties;

        foreach (var r in firstPara.Elements<W.Run>().ToList())
            r.Remove();

        var newRun = new W.Run();
        if (rpr != null) newRun.Append(rpr);
        newRun.Append(new W.Text(rocDate) { Space = SpaceProcessingModeValues.Preserve });
        firstPara.Append(newRun);

        foreach (var p in paragraphs.Skip(1).ToList())
            p.Remove();
    }

    // ─── 2. Replace "日期:XXX" in body paragraphs ──────────────────────────────

    private static readonly Regex DateLabelRegex =
        new(@"(日期[：:])(.+)", RegexOptions.Compiled);

    private static void ReplaceInlineBodyDates(W.Body body, string rocDate)
    {
        foreach (var para in body.Descendants<W.Paragraph>())
            ReplaceDateInParagraph(para, rocDate);
    }

    private static void ReplaceDateInParagraph(W.Paragraph para, string rocDate)
    {
        var runs = para.Elements<W.Run>().ToList();
        if (runs.Count == 0) return;

        var fullText = string.Concat(runs.Select(r =>
            string.Concat(r.Elements<W.Text>().Select(t => t.Text))));

        var match = DateLabelRegex.Match(fullText);
        if (!match.Success) return;

        // Preserve text before "日期:" (e.g. "查驗") + label + new date
        var newFullText = fullText.Substring(0, match.Index) + match.Groups[1].Value + rocDate;

        var firstRun = runs[0];
        foreach (var t in firstRun.Elements<W.Text>().ToList())
            t.Remove();
        firstRun.Append(new W.Text(newFullText) { Space = SpaceProcessingModeValues.Preserve });

        foreach (var run in runs.Skip(1))
            foreach (var t in run.Elements<W.Text>().ToList())
                t.Remove();
    }
}
