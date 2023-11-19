using ClosedXML.Excel;
using System.Reflection;

var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Sheet1");


var Departments = new List<string>
                    {
                        "Finance",
                        "Human Resources",
                        "Marketing",
                        "Information Technology",
                        "Operations"
                    };

var Countries = new List<string>
        {
            "United States",
            "Canada",
            "United Kingdom",
            "Germany",
            "France"
        };


//Heading
worksheet.Cell("A1").Value = "Country";
worksheet.Cell("B1").Value = "Name";
worksheet.Cell("C1").Value = "Department";



//Validation
//Dropdown_Department
var validOptions = $"\"{String.Join(",", Departments)}\"";
for (int i = 2; i <= 500; i++)
{
    worksheet.Cell(i, 1).GetDataValidation().List(validOptions, true);
}


//Dropdown_Country
validOptions = $"\"{String.Join(",", Countries)}\"";
for (int i = 2; i <= 500; i++)
{
    worksheet.Cell(i, 3).GetDataValidation().List(validOptions, true);
}




//Styles
worksheet.Range(1, 1, 1, 3).Style.Fill.BackgroundColor = XLColor.Yellow;
var rangeWithBorders = worksheet.Range("A1:C500");

// Apply border formatting
rangeWithBorders.Style.Border.TopBorder = XLBorderStyleValues.Thin;
rangeWithBorders.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
rangeWithBorders.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
rangeWithBorders.Style.Border.RightBorder = XLBorderStyleValues.Thin;


//Download
var folderPath = Directory.GetCurrentDirectory().Replace(@"bin\Debug\net7.0", "") + "Files\\";
var fileName = folderPath + "ExcelWithDropdown_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + ".xlsx";

workbook.SaveAs(fileName);

Console.WriteLine("Excel file with dropdown created successfully.");
Console.WriteLine("File name: " + fileName);