using System.Runtime.InteropServices;

namespace WordAutoTool.Services;

/// <summary>
/// All operations requiring Word COM Automation.
/// Uses Type.GetTypeFromProgID + dynamic to avoid requiring Office PIAs at runtime.
/// Word must be installed on the target machine.
/// </summary>
public static class WordComService
{
    // Word enum constants (avoids needing the interop assembly)
    private const int WdFormatDocument    = 0;   // wdFormatDocument (.doc 97-2003)
    private const int WdFormatXMLDocument = 12;  // wdFormatXMLDocument (.docx)
    private const int WdAlertsNone        = 0;   // wdAlertsNone
    private const int WdDoNotSaveChanges  = 0;   // wdDoNotSaveChanges

    // ── Format detection ──────────────────────────────────────────────────────

    private static readonly byte[] DocHeader = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    public static bool IsLegacyDoc(byte[] bytes)
    {
        if (bytes.Length < DocHeader.Length) return false;
        for (int i = 0; i < DocHeader.Length; i++)
            if (bytes[i] != DocHeader[i]) return false;
        return true;
    }

    // ── .docx → .doc conversion ───────────────────────────────────────────────

    public static byte[] ConvertDocxToDoc(byte[] docxBytes)
    {
        var tempDocx = TempFile(".docx");
        var tempDoc  = TempFile(".doc");
        try
        {
            File.WriteAllBytes(tempDocx, docxBytes);
            WithWord(word =>
            {
                dynamic doc = word.Documents.Open(
                    tempDocx,  // FileName
                    false,     // ConfirmConversions
                    true,      // ReadOnly
                    false);    // AddToRecentFiles
                doc.SaveAs2(tempDoc, WdFormatDocument);
                doc.Close(WdDoNotSaveChanges);
                Marshal.ReleaseComObject(doc);
            });
            return File.ReadAllBytes(tempDoc);
        }
        finally { TryDelete(tempDocx); TryDelete(tempDoc); }
    }

    // ── .doc → .docx conversion ───────────────────────────────────────────────

    public static byte[] ConvertDocToDocx(byte[] docBytes)
    {
        var tempDoc  = TempFile(".doc");
        var tempDocx = TempFile(".docx");
        try
        {
            File.WriteAllBytes(tempDoc, docBytes);
            WithWord(word =>
            {
                dynamic doc = word.Documents.Open(
                    tempDoc,   // FileName
                    false,     // ConfirmConversions
                    true,      // ReadOnly
                    false);    // AddToRecentFiles
                doc.SaveAs2(tempDocx, WdFormatXMLDocument);
                doc.Close(WdDoNotSaveChanges);
                Marshal.ReleaseComObject(doc);
            });
            return File.ReadAllBytes(tempDocx);
        }
        finally { TryDelete(tempDoc); TryDelete(tempDocx); }
    }

    // ── Image replacement ─────────────────────────────────────────────────────

    public static (byte[] DocxBytes, List<string> Log) ReplaceImages(
        byte[] docxBytes,
        List<(string Name, byte[] Data, string ContentType)> images)
    {
        if (images.Count == 0) return (docxBytes, []);

        var tempDocx   = TempFile(".docx");
        var tempImages = new List<string>();
        var log        = new List<string>();
        try
        {
            File.WriteAllBytes(tempDocx, docxBytes);

            foreach (var (_, data, contentType) in images)
            {
                var tf = TempFile(ContentTypeToExt(contentType, data));
                File.WriteAllBytes(tf, data);
                tempImages.Add(tf);
            }

            WithWord(word =>
            {
                dynamic doc = word.Documents.Open(
                    tempDocx,  // FileName
                    false,     // ConfirmConversions
                    false,     // ReadOnly
                    false);    // AddToRecentFiles

                dynamic inlineShapes = doc.InlineShapes;
                int count = (int)inlineShapes.Count;

                // Pre-collect dimensions (1-based array)
                var dims = new (float w, float h)[count + 1];
                for (int i = 1; i <= count; i++)
                {
                    dynamic s = inlineShapes[i];
                    dims[i] = ((float)s.Width, (float)s.Height);
                    Marshal.ReleaseComObject(s);
                }

                int replaceCount = Math.Min(count, tempImages.Count);

                // Process in reverse so deleting shape[i] doesn't shift earlier indices
                for (int i = replaceCount; i >= 1; i--)
                {
                    dynamic shape = inlineShapes[i];
                    shape.Select();
                    Marshal.ReleaseComObject(shape);

                    word.Selection.Delete();
                    word.Selection.InlineShapes.AddPicture(tempImages[i - 1]);

                    dynamic newShape = inlineShapes[i];
                    newShape.Width  = dims[i].w;
                    newShape.Height = dims[i].h;
                    Marshal.ReleaseComObject(newShape);
                }

                Marshal.ReleaseComObject(inlineShapes);

                // Build log in forward order
                for (int i = 1; i <= replaceCount; i++)
                    log.Add($"✅ 圖片 {i}：已替換");
                if (tempImages.Count > count)
                    log.Add($"⚠ 圖片 {count + 1}～{tempImages.Count}：文件中只有 {count} 張內嵌圖片，多餘圖片略過");
                else if (count > tempImages.Count)
                    log.Add($"（文件共 {count} 張內嵌圖片，本次替換了前 {tempImages.Count} 張）");

                doc.Save();
                doc.Close(WdDoNotSaveChanges);
                Marshal.ReleaseComObject(doc);
            });

            return (File.ReadAllBytes(tempDocx), log);
        }
        finally
        {
            TryDelete(tempDocx);
            foreach (var f in tempImages) TryDelete(f);
        }
    }

    // ── Word lifetime ─────────────────────────────────────────────────────────

    private static void WithWord(Action<dynamic> action)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("找不到 Word.Application，請確認已安裝 Microsoft Word。");

        dynamic word = Activator.CreateInstance(wordType)!;
        word.Visible       = false;
        word.DisplayAlerts = WdAlertsNone;

        try   { action(word); }
        finally
        {
            try { word.Quit(WdDoNotSaveChanges); } catch { }
            Marshal.ReleaseComObject(word);
        }
    }

    // ── Utilities ─────────────────────────────────────────────────────────────

    private static string TempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"wt_{Guid.NewGuid():N}{ext}");

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    private static string ContentTypeToExt(string contentType, byte[] data)
    {
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8) return ".jpg";
        if (data.Length >= 4 && data[0] == 0x89 && data[1] == 0x50) return ".png";
        if (data.Length >= 4 && data[0] == 0x47 && data[1] == 0x49) return ".gif";
        if (data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D) return ".bmp";
        return contentType switch
        {
            "image/jpeg" => ".jpg",
            "image/png"  => ".png",
            "image/gif"  => ".gif",
            "image/bmp"  => ".bmp",
            _            => ".png"
        };
    }
}
