namespace MyUtilities.Models;

public class SiteExcelModel
{
    public string OrderNumber { get; set; }

    //نام محصولی که سفارش داده
    public string Name { get; set; }
    public string FullName { get; set; }
    public string StateName { get; set; }
    public string Address { get; set; }
    public string? Email { get; set; }
    public string? Phone { get; set; }
    public string? Zip { get; set; }
    public string? Description { get; set; }
    public decimal ItemCoast { get; set; }
    public bool? PostCoast { get; set; }
    public int Quantity { get; set; }
    public string City { get; set; }
    public DateTime OrderDate { get; set; }
    public string ShamsiDate { get; set; }


}
