using ClosedXML.Excel;
using MyUtilities.Models;
using System.Globalization;
using System.Net.Mime;

namespace MyUtilities.Helpers;

public static class ExcelServices
{
    private static readonly PersianCalendar persianCalendar = new PersianCalendar();

    public static List<SiteExcelModel> ReadSiteExcel(IFormFile siteExcel)
    {
        try
        {
            var fileStream = new MemoryStream();
            siteExcel.CopyTo(fileStream);
            fileStream.Position = 0;
            var list = new List<SiteExcelModel>();
            using var workbook = new XLWorkbook(fileStream);
            var worksheet = workbook.Worksheet(1);
            int row = 2;// از ردیف 2 شروع میکنیم چون ردیف یک هدره
            while (!worksheet.Cell(row, 1).IsEmpty())
            {
                var item = new SiteExcelModel();

                item.OrderNumber = worksheet.Cell(row, 1).GetValue<string>();     // A
                item.Name = worksheet.Cell(row, 2).GetValue<string>();     // B
                item.FullName = worksheet.Cell(row, 3).GetValue<string>();  // C
                item.StateName = worksheet.Cell(row, 4).GetValue<string>();      // D
                item.Address = worksheet.Cell(row, 5).GetValue<string>();// E
                item.Email = worksheet.Cell(row, 6).GetValue<string>();     // F
                item.Phone = worksheet.Cell(row, 7).GetValue<string>();     // G
                item.Zip = worksheet.Cell(row, 8).GetValue<string>();     // H
                item.Description = worksheet.Cell(row, 9).GetValue<string>();     // I
                item.ItemCoast = worksheet.Cell(row, 10).GetValue<decimal>();// J
                item.PostCoast = worksheet.Cell(row, 11).GetValue<bool>();//K
                item.Quantity = worksheet.Cell(row, 14).GetValue<int>(); // N
                item.City = worksheet.Cell(row, 15).GetValue<string>();// O
                var datestring = worksheet.Cell(row, 16).GetValue<string>();
                item.OrderDate = DateTime.Parse( datestring) ; // P
                item.ShamsiDate = item.OrderDate.ToShamsi();
                list.Add(item);
                row++;
            }
            return list;
        }

        catch(Exception ex) 
        {
            throw new Exception(ex.Message);
        }
    }

    public static WarehouseFileDto SetWarehouseFile(IFormFile warehouseFile, List<SiteExcelModel> items)
    {
        var fileStream = new MemoryStream();
        warehouseFile.CopyTo(fileStream);
        fileStream.Position = 0;
        using var workbook = new XLWorkbook(fileStream);
        var worksheet = workbook.Worksheet(1);
        var fileName = warehouseFile.FileName;
        var contentType = warehouseFile.ContentType;
        // پیدا کردن آخرین ردیف استفاده شده
        int lastRow = worksheet.LastRowUsed().RowNumber();
        int row = lastRow + 1;
        foreach (var item in items)
        {
            worksheet.Cell(row, 2).Value = item.OrderNumber; // B
            worksheet.Cell(row, 3).Value = $"{item.Name} (سایت)";        // C
            worksheet.Cell(row, 4).Value = item.FullName;    // D
            worksheet.Cell(row, 5).Value = item.StateName;   // E
            worksheet.Cell(row, 6).Value = $"{item.City ?? ""}، {item.Address ?? ""}";// F
            worksheet.Cell(row, 7).Value = item.Email;       // G
            worksheet.Cell(row, 8).Value = item.Phone;       // H
            worksheet.Cell(row, 9).Value = item.Zip;       // I
            worksheet.Cell(row, 10).Value = item.Description;       // J
            worksheet.Cell(row, 11).Value = item.ItemCoast/1000;   // K
            //worksheet.Cell(row, 12).Value = item.PostCoast;   // L
            worksheet.Cell(row, 12).Value = item.PostCoast == true ? 1 : 0;
            worksheet.Cell(row, 13).Value = item.ShamsiDate;   // M
            worksheet.Cell(row, 14).Value = item.OrderNumber;  // o

            row++;
        }

        using var outputStream = new MemoryStream();
        workbook.SaveAs(outputStream);

        return new WarehouseFileDto
        {
            FileByte = outputStream.ToArray(),
            FileName = fileName,
            ContentType = contentType
        };
    }

    private static bool IsExcelFile(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLower();

        return extension == ".xlsx" ||
               extension == ".xls" ||
               extension == ".xlsm";
    }

    public static string ToShamsi(this DateTime date)
    {

        return $"{persianCalendar.GetYear(date)}/{persianCalendar.GetMonth(date):D2}/{persianCalendar.GetDayOfMonth(date):D2}";
    }

}
