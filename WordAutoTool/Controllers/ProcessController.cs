using System.Drawing;
using System.IO.Compression;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using WordAutoTool.Services;

namespace WordAutoTool.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ProcessController : ControllerBase
{
    private readonly WordProcessingService _wordService;

    public ProcessController(WordProcessingService wordService)
    {
        _wordService = wordService;
    }

    private static byte[] EnsureLandscape(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var img = Image.FromStream(ms);
        if (img.Height <= img.Width) return data;
        using var bmp = new Bitmap(img);
        bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
        using var outMs = new MemoryStream();
        bmp.Save(outMs, img.RawFormat);
        return outMs.ToArray();
    }

    [HttpPost]
    [RequestSizeLimit(500 * 1024 * 1024)]
    public async Task<IActionResult> Process(
        [FromForm] IFormFile wordFile,
        [FromForm] IFormFileCollection images,
        [FromForm] string date,
        [FromForm] string? padZero,
        [FromForm] string? includePdf,
        [FromForm] string? overwrite)
    {
        if (wordFile == null || wordFile.Length == 0)
            return BadRequest(new { error = "請選擇 Word 檔案" });

        if (string.IsNullOrWhiteSpace(date))
            return BadRequest(new { error = "請選擇日期" });

        if (!DateOnly.TryParse(date, out var parsedDate))
            return BadRequest(new { error = "日期格式無效" });

        bool pad = padZero == "true";
        int rocYear = parsedDate.Year - 1911;
        // Word 內文日期：依選項決定是否補零（例：115.3.5 或 115.03.05）
        string rocDate = pad
            ? $"{rocYear}.{parsedDate.Month:D2}.{parsedDate.Day:D2}"
            : $"{rocYear}.{parsedDate.Month}.{parsedDate.Day}";
        // 檔名用月日各兩位（例：0329）
        string fileMonthDay = $"{parsedDate.Month:D2}{parsedDate.Day:D2}";

        var log = new List<string>();

        // Read uploaded file
        byte[] fileBytes;
        using (var ms = new MemoryStream())
        {
            await wordFile.CopyToAsync(ms);
            fileBytes = ms.ToArray();
        }

        if (fileBytes.Length < 8)
            return BadRequest(new { error = "檔案內容無效（太小）" });

        // Step 1: Convert .doc → .docx if needed
        byte[] docxBytes;
        if (WordComService.IsLegacyDoc(fileBytes))
        {
            try
            {
                docxBytes = WordComService.ConvertDocToDocx(fileBytes);
                log.Add("✅ .doc → .docx 轉換完成");
            }
            catch (Exception ex)
            { return StatusCode(500, new { error = $".doc 轉換失敗：{ex.Message}" }); }
        }
        else if (fileBytes[0] == 0x50 && fileBytes[1] == 0x4B)
        {
            docxBytes = fileBytes;
        }
        else
        {
            return BadRequest(new { error = "不支援的格式，請上傳 .doc 或 .docx" });
        }

        // Step 2: Replace images via Word COM (preserves position & size)
        var imageList = new List<(string Name, byte[] Data, string ContentType)>();
        foreach (var img in images
            .Where(f => f.Length > 0)
            .OrderBy(f => f.FileName, StringComparer.OrdinalIgnoreCase))
        {
            using var imgMs = new MemoryStream();
            await img.CopyToAsync(imgMs);
            imageList.Add((img.FileName, imgMs.ToArray(), img.ContentType));
        }

        // Rotate portrait images to landscape before replacement
        for (int i = 0; i < imageList.Count; i++)
        {
            var (name, data, ct) = imageList[i];
            imageList[i] = (name, EnsureLandscape(data), ct);
        }

        if (imageList.Count > 0)
        {
            try
            {
                var (newDocx, imgLog) = WordComService.ReplaceImages(docxBytes, imageList);
                docxBytes = newDocx;
                log.AddRange(imgLog);
            }
            catch (Exception ex)
            { return StatusCode(500, new { error = $"圖片替換失敗：{ex.Message}" }); }
        }
        else
        {
            log.Add("（未上傳圖片，略過圖片替換）");
        }

        // Step 3: Replace text boxes + body dates via OpenXML
        byte[] docxResult;
        try
        {
            docxResult = _wordService.Process(docxBytes, rocDate);
            log.Add($"✅ 日期已更新為 {rocDate}");
        }
        catch (Exception ex)
        { return StatusCode(500, new { error = $"文字處理失敗：{ex.Message}" }); }

        // Step 4: Convert .docx → .doc
        byte[] resultBytes;
        try
        {
            resultBytes = WordComService.ConvertDocxToDoc(docxResult);
            log.Add("✅ 已轉換為 .doc 格式");
        }
        catch (Exception ex)
        { return StatusCode(500, new { error = $".doc 轉換失敗：{ex.Message}" }); }

        // Attach log as response header (URI-encoded JSON array)
        Response.Headers.Append("X-Process-Log", Uri.EscapeDataString(JsonSerializer.Serialize(log)));

        // Overwrite mode: return .doc with the original filename (no rename, no PDF)
        if (overwrite == "true")
        {
            string origName = Path.GetFileNameWithoutExtension(wordFile.FileName) + ".doc";
            Response.Headers["Content-Disposition"] =
                $"attachment; filename*=UTF-8''{Uri.EscapeDataString(origName)}";
            return File(resultBytes, "application/msword");
        }

        string baseName = $"8_查驗照片{fileMonthDay}";

        if (includePdf == "true")
        {
            // Convert .doc → PDF, then zip both
            byte[] pdfBytes;
            try
            {
                pdfBytes = WordComService.ConvertDocToPdf(resultBytes);
                log.Add("✅ 已轉換為 PDF 格式");
            }
            catch (Exception ex)
            { return StatusCode(500, new { error = $"PDF 轉換失敗：{ex.Message}" }); }

            // Re-write log header now that PDF entry is appended
            Response.Headers["X-Process-Log"] = Uri.EscapeDataString(JsonSerializer.Serialize(log));

            using var zipMs = new MemoryStream();
            using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var docEntry = archive.CreateEntry($"{baseName}.doc", CompressionLevel.Fastest);
                using (var s = docEntry.Open()) s.Write(resultBytes);

                var pdfEntry = archive.CreateEntry($"{baseName}.pdf", CompressionLevel.Fastest);
                using (var s = pdfEntry.Open()) s.Write(pdfBytes);
            }

            string zipName = $"{baseName}.zip";
            Response.Headers["Content-Disposition"] =
                $"attachment; filename*=UTF-8''{Uri.EscapeDataString(zipName)}";
            return File(zipMs.ToArray(), "application/zip");
        }
        else
        {
            string fileName = $"{baseName}.doc";
            Response.Headers["Content-Disposition"] =
                $"attachment; filename*=UTF-8''{Uri.EscapeDataString(fileName)}";
            return File(resultBytes, "application/msword");
        }
    }
}
