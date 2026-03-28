using Microsoft.AspNetCore.Mvc;
using WordAutoTool.Services;

namespace WordAutoTool.Controllers;

[ApiController]
[Route("api/[controller]")]
public class InspectController : ControllerBase
{
    private readonly InspectService _inspectService;

    public InspectController(InspectService inspectService)
    {
        _inspectService = inspectService;
    }

    [HttpPost]
    [RequestSizeLimit(500 * 1024 * 1024)]
    public async Task<IActionResult> Inspect([FromForm] IFormFile wordFile)
    {
        if (wordFile == null || wordFile.Length == 0)
            return BadRequest(new { error = "請提供檔案" });

        byte[] fileBytes;
        using (var ms = new MemoryStream())
        {
            await wordFile.CopyToAsync(ms);
            fileBytes = ms.ToArray();
        }

        // Legacy .doc: report it but can't inspect without COM conversion
        bool isLegacy = WordComService.IsLegacyDoc(fileBytes);
        if (isLegacy)
            return Ok(new { isLegacy = true });

        // Not a ZIP/docx either
        if (fileBytes.Length < 4 || fileBytes[0] != 0x50 || fileBytes[1] != 0x4B)
            return BadRequest(new { error = "不支援的格式" });

        var result = _inspectService.Inspect(fileBytes);
        return Ok(new
        {
            isLegacy       = false,
            floatingShapes = result.FloatingShapes.Select(s => new { s.Name, s.Type, s.Text }),
            inlineShapes   = result.InlineShapes.Select(s => new { s.Name, s.Type }),
            textBoxes      = result.TextBoxes.Select(s => new { s.Name, s.Type, s.Text }),
            bodyParagraphs = result.BodyParagraphs,
        });
    }
}
