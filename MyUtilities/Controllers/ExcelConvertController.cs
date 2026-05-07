using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MyUtilities.Helpers;
using MyUtilities.Models;

namespace MyUtilities.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ExcelConvertController : ControllerBase
{

    [HttpPost]
    public async Task<IActionResult> CovretSiteExcel([FromForm]FilesDto files)
    {
        var siteExcelList = ExcelServices.ReadSiteExcel(files.SiteExcel);

        var warehouseFile = ExcelServices.SetWarehouseFile(files.WarehouseExcel, siteExcelList);

        return File(warehouseFile.FileByte, warehouseFile.ContentType, warehouseFile.FileName);
    }
}
