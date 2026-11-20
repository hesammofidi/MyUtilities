using System.ComponentModel.DataAnnotations;

namespace MyUtilities.Models;

public class WarehouseSaleFileDto
{
    [Required]
    public IFormFile WarehouseExcel { get; set; } = null!;

    [Required(ErrorMessage = "تاریخ شروع را به میلادی با فرمت yyyy-MM-dd وارد کنید.")]
    public DateOnly? StartDate { get; set; }
}
