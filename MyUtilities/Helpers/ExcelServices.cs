using ClosedXML.Excel;
using MyUtilities.Models;
using System.Globalization;
using System.Net.Mime;

namespace MyUtilities.Helpers;

public static class ExcelServices
{
    private static readonly PersianCalendar persianCalendar = new PersianCalendar();

    public static WarehouseFileDto CreateSaleProductCountExcel(IFormFile siteExcel)
    {
        var productCounts = new Dictionary<string, int>();
        using (var inputStream = siteExcel.OpenReadStream())
        using (var sourceWorkbook = new XLWorkbook(inputStream))
        {
            var worksheet = sourceWorkbook.Worksheet(1);
            foreach (var cell in worksheet.Column(2).CellsUsed(XLCellsUsedOptions.Contents))
            {
                if (cell.Address.RowNumber == 1)
                    continue;

                var productName = cell.GetString();
                if (string.IsNullOrWhiteSpace(productName))
                    continue;

                productCounts.TryGetValue(productName, out var count);
                productCounts[productName] = count + 1;
            }
        }

        return CreateProductCountWorkbook(productCounts);
    }

    public static WarehouseFileDto CreateWarehouseSaleProductCountExcel(IFormFile warehouseExcel, DateOnly startDate)
    {
        string shamsiStartDate;
        try
        {
            shamsiStartDate = startDate.ToDateTime(TimeOnly.MinValue).ToShamsi();
        }
        catch (ArgumentOutOfRangeException)
        {
            throw new FormatException("تاریخ شروع خارج از محدوده تقویم شمسی است.");
        }

        if (!TryParseShamsiDate(shamsiStartDate, out var minimumDate))
            throw new FormatException("تاریخ شروع قابل تبدیل به تاریخ شمسی نیست.");
        var productCounts = new Dictionary<string, int>();
        using (var inputStream = warehouseExcel.OpenReadStream())
        using (var sourceWorkbook = new XLWorkbook(inputStream))
        {
            var worksheet = sourceWorkbook.Worksheet(1);
            foreach (var cell in worksheet.Column(3).CellsUsed(XLCellsUsedOptions.Contents))
            {
                var row = cell.Address.RowNumber;
                if (row == 1)
                    continue;

                var productName = cell.GetString();
                if (string.IsNullOrWhiteSpace(productName))
                    continue;

                var dateCell = worksheet.Cell(row, 13);
                DateTime orderDate;
                if (dateCell.DataType == XLDataType.DateTime)
                    orderDate = dateCell.GetDateTime().Date;
                else if (!TryParseShamsiDate(dateCell.GetString(), out orderDate))
                    continue;

                if (orderDate < minimumDate)
                    continue;

                productCounts.TryGetValue(productName, out var count);
                productCounts[productName] = count + 1;
            }
        }

        return CreateProductCountWorkbook(productCounts);
    }

    private static bool TryParseShamsiDate(string value, out DateTime date)
    {
        date = default;
        value = new string((value ?? string.Empty)
            .Where(character => char.GetUnicodeCategory(character) != UnicodeCategory.Format)
            .Select(character => character switch
            {
                >= '\u06F0' and <= '\u06F9' => (char)('0' + character - '\u06F0'),
                >= '\u0660' and <= '\u0669' => (char)('0' + character - '\u0660'),
                _ => character
            }).ToArray()).Trim();

        var parts = value.Split('/', StringSplitOptions.TrimEntries);
        if (parts.Length != 3 || parts[0].Length < 1 || parts[0].Length > 4 ||
            parts[1].Length < 1 || parts[1].Length > 2 || parts[2].Length < 1 || parts[2].Length > 2 ||
            !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var year) ||
            !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var month) ||
            !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var day))
            return false;

        try
        {
            date = persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private static WarehouseFileDto CreateProductCountWorkbook(Dictionary<string, int> productCounts)
    {
        using var workbook = new XLWorkbook();
        var report = workbook.Worksheets.Add("تعداد فروش محصولات");
        report.RightToLeft = true;
        report.Cell(1, 1).Value = "نام محصول";
        report.Cell(1, 2).Value = "تعداد فروش";

        var row = 2;
        foreach (var product in productCounts)
        {
            report.Cell(row, 1).Value = product.Key;
            report.Cell(row, 2).Value = product.Value;
            row++;
        }

        var usedRange = report.Range(1, 1, row - 1, 2);
        usedRange.Style.Font.FontName = "Tahoma";
        var header = report.Range(1, 1, 1, 2);
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.LightBlue;
        report.Column(1).Width = 60;
        report.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        usedRange.Style.Alignment.WrapText = true;
        report.Column(2).Width = 18;
        report.Column(2).Style.NumberFormat.Format = "#,##0";
        report.SheetView.FreezeRows(1);

        using var outputStream = new MemoryStream();
        workbook.SaveAs(outputStream);
        return new WarehouseFileDto
        {
            FileByte = outputStream.ToArray(),
            FileName = "گزارش تعداد فروش محصولات.xlsx",
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };
    }

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
