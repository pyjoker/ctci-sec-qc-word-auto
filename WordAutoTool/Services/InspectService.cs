using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW  = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace WordAutoTool.Services;

public class InspectService
{
    public InspectResult Inspect(byte[] docxBytes)
    {
        var result = new InspectResult();
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);

        var body = doc.MainDocumentPart!.Document.Body!;

        // ── 浮動圖形（Floating shapes via wp:anchor）────────────────────────────
        foreach (var anchor in body.Descendants<DW.Anchor>())
        {
            var docPr   = anchor.GetFirstChild<DW.DocProperties>();
            string name = docPr?.Name?.Value ?? "（無名稱）";

            bool hasPic     = anchor.Descendants<PIC.Picture>().Any();
            bool hasTxbx    = anchor.Descendants<WPS.TextBoxInfo2>().Any();
            string typeLabel = hasPic ? "圖片" : hasTxbx ? "文字方塊" : "圖形";

            string text = "";
            if (hasTxbx)
            {
                var content = anchor.Descendants<TextBoxContent>().FirstOrDefault();
                if (content != null)
                    text = string.Concat(content.Descendants<Text>().Select(t => t.Text)).Trim();
            }

            result.FloatingShapes.Add(new ShapeInfo(name, typeLabel, text));
        }

        // ── 浮動圖形（Legacy VML via w:pict / v:shape）──────────────────────────
        foreach (var pict in body.Descendants<Picture>())
        {
            foreach (var vShape in pict.ChildElements
                .Where(e => e.LocalName == "shape"))
            {
                string id   = vShape.GetAttribute("id",   "").Value;
                string name = vShape.GetAttribute("alt",  "").Value;
                if (string.IsNullOrEmpty(name)) name = id;
                if (string.IsNullOrEmpty(name)) name = "（無名稱）";

                result.FloatingShapes.Add(new ShapeInfo(name, "圖片(VML)", ""));
            }
        }

        // ── 內嵌圖形（Inline shapes via wp:inline）──────────────────────────────
        foreach (var inline in body.Descendants<DW.Inline>())
        {
            var docPr   = inline.GetFirstChild<DW.DocProperties>();
            string name = docPr?.Name?.Value ?? "（無名稱）";
            result.InlineShapes.Add(new ShapeInfo(name, "內嵌圖片", ""));
        }

        // ── 文字方塊（w:txbxContent，涵蓋 VML 文字方塊）────────────────────────
        foreach (var txbx in body.Descendants<TextBoxContent>())
        {
            // 已在浮動圖形裡收過 DrawingML 文字方塊，這裡只補 VML 文字方塊
            // 判斷：父鏈不含 DW.Anchor
            bool underAnchor = txbx.Ancestors<DW.Anchor>().Any();
            if (underAnchor) continue;

            string text = string.Concat(txbx.Descendants<Text>().Select(t => t.Text)).Trim();

            // 試著找上層 v:shape 的名稱
            var vShape   = txbx.Ancestors().FirstOrDefault(e => e.LocalName == "shape");
            string vId   = vShape?.GetAttribute("id",  "").Value ?? "";
            string title = vShape?.GetAttribute("alt", "").Value ?? "";
            string label = !string.IsNullOrEmpty(title) ? title
                         : !string.IsNullOrEmpty(vId)   ? $"(id={vId})"
                         : "（無名稱）";

            result.TextBoxes.Add(new ShapeInfo(label, "VML文字方塊", text));
        }

        // ── 段落文字（Body paragraphs）──────────────────────────────────────────
        foreach (var para in body.Elements<Paragraph>())
        {
            var runs = para.Elements<Run>().ToList();
            if (runs.Count == 0) continue;

            string text = string.Concat(runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text)).Trim();
            if (text.Length == 0) continue;

            result.BodyParagraphs.Add(text);
        }

        return result;
    }
}

public record ShapeInfo(string Name, string Type, string Text);

public class InspectResult
{
    public List<ShapeInfo> FloatingShapes   { get; } = [];
    public List<ShapeInfo> InlineShapes     { get; } = [];
    public List<ShapeInfo> TextBoxes        { get; } = [];
    public List<string>    BodyParagraphs   { get; } = [];
}
