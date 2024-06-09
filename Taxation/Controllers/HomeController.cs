using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Text;
using Taxation.Models;

namespace Taxation.Controllers
{
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;

		public HomeController(ILogger<HomeController> logger)
		{
			_logger = logger;
		}

		public IActionResult Index()
		{
			return View();
		}

		public IActionResult Privacy()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}
		[HttpGet]
		public IActionResult ExcelFileReader()
		{
			return View();
		}
		[HttpPost]
		public async Task<IActionResult> ExcelFileReader(IFormFile file)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			//Upload File
			if (file is not null && file.Length > 0)
			{
				var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads";
                if (!Directory.Exists(uploadDirectory))
                {
					Directory.CreateDirectory(uploadDirectory);
                }
				var filePath = Path.Combine(uploadDirectory, file.Name);
				using (var stream = new FileStream(filePath , FileMode.Create))
				{
					await file.CopyToAsync(stream);
				}
                // ReadFile 
                var excelData = new List<List<object>>();
                double sumTotalBeforeTaxes = 0;
                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
					
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
						bool IsSpecRow = true;


						var numberOfRows = reader.RowCount;
						var Fcount = reader.FieldCount;
						ViewBag.Fcount = Fcount;
						ViewBag.NumberOfRows = numberOfRows;
                        do
                        {
                            while (reader.Read())
                            {
								var rowData = new List<object>();
								double totalAfetrtaxing = 0, taxes = 0 , totalBeforTaxes =0 ;

								for(int column = 0; column < reader.FieldCount; column++)
								{
									var cellValue = reader.GetValue(column);
									rowData.Add(cellValue);
                                    if ( column == 6 && double.TryParse(cellValue?.ToString(),out double temptotal))
                                    {
                                        totalAfetrtaxing = temptotal;
                                    }
                                    if (column == 7 && double.TryParse(cellValue?.ToString(), out double temptaxes))
                                    {
                                        taxes = temptaxes;
                                    }

                                }

                                if (IsSpecRow)
                                {
									rowData.Add("Total Befor taxing");
									IsSpecRow = false;
                                    
                                }
								else
								{
									
									totalBeforTaxes = totalAfetrtaxing - taxes;
                                    sumTotalBeforeTaxes += totalBeforTaxes;
                                    rowData.Add( totalBeforTaxes);
									
								}
                                excelData.Add(rowData);
                            }
                        } while (reader.NextResult());
                        var sumRow = new List<object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            sumRow.Add(i == reader.FieldCount - 1 ? "Total Sum" : "");
                        }
						
                        ViewBag.excelData = excelData;
                        ViewBag.SumTotalBeforeTaxes = sumTotalBeforeTaxes;

                    }
                }



            }
			return View();
		}

        public FileResult ExportToExcel(string htmlTable)
        {
            try
            {
                if (string.IsNullOrEmpty(htmlTable))
                {
                    throw new ArgumentException("The HTML table content is empty or null.");
                }

                byte[] fileContent = Encoding.ASCII.GetBytes(htmlTable);

                return File(fileContent, "application/vnd.ms-excel", "Taxes_Sheet.xls");
            }
            catch (Exception ex)
            {
                return File(Encoding.ASCII.GetBytes("An error occurred while generating the Excel file."), "application/vnd.ms-excel", "Error.xls");
            }
        }
    }
}
