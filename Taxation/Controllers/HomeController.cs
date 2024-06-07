using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
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
                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
					
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
						bool IsSpecRow = true;


						var numberOfRows = reader.RowCount;
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
									rowData.Add( totalBeforTaxes);
									
								}
                                excelData.Add(rowData);
                            }
                        } while (reader.NextResult());
						ViewBag.excelData = excelData;

                    }
                }



            }
			return View();
		}
	}
}
