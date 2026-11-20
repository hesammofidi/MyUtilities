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

    [HttpPost("sale-product-count-excel")]
    public IActionResult SaleProductCountExcel([FromForm] SaleFileDto files)
    {
        var report = ExcelServices.CreateSaleProductCountExcel(files.SiteExcel);
        return File(report.FileByte, report.ContentType, report.FileName);
    }

    [HttpPost("warehouse-sale-product-count-excel")]
    public IActionResult WarehouseSaleProductCountExcel([FromForm] WarehouseSaleFileDto files)
    {
        if (!files.StartDate.HasValue)
            return BadRequest(new { message = "تاریخ شروع را به میلادی با فرمت yyyy-MM-dd وارد کنید." });

        try
        {
            var report = ExcelServices.CreateWarehouseSaleProductCountExcel(files.WarehouseExcel, files.StartDate.Value);
            return File(report.FileByte!, report.ContentType!, report.FileName);
        }
        catch (FormatException ex)
        {
            return BadRequest(new { message = ex.Message });
        }
    }

    [HttpPost("warehouse-sale-product-count")]
    [ProducesResponseType(typeof(List<SaleReportByProductDto>), StatusCodes.Status200OK)]
    public IActionResult WarehouseSaleProductCount([FromForm] WarehouseSaleFileDto files)
    {
        if (!files.StartDate.HasValue)
            return BadRequest(new { message = "تاریخ شروع را به میلادی با فرمت yyyy-MM-dd وارد کنید." });

        try
        {
            var productCounts = ExcelServices.GetWarehouseSaleProductCounts(files.WarehouseExcel, files.StartDate.Value);
            var report = productCounts.Select(product => new SaleReportByProductDto
            {
                Name = product.Key,
                Count = product.Value
            }).ToList();
            return Ok(report);
        }
        catch (FormatException ex)
        {
            return BadRequest(new { message = ex.Message });
        }
    }

    [HttpPost("sale-product-count")]
    public async Task<IActionResult> SaleProductCount([FromForm] SaleFileDto files)
    {
        var siteExcelList = ExcelServices.ReadSiteExcel(files.SiteExcel);
        var siteExcelListGroupByName = siteExcelList.GroupBy(x => x.Name);
        var response = new List<SaleReportByProductDto>();
        foreach(var item in siteExcelListGroupByName)
        {
            var responseItem = new SaleReportByProductDto();
            responseItem.Name = item.Key;
            responseItem.Count = item.Count();
            response.Add(responseItem);
        }
        return Ok(response);
    }
}
public class SaleFileDto
{
    public IFormFile SiteExcel { get; set; }
}

public class SaleReportByProductDto 
{
    public string Name { get; set; }
    public int Count { get; set; }
}
