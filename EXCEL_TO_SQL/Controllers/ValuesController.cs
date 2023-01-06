using Dapper;
using EXCEL_TO_SQL.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data.SqlClient;

namespace EXCEL_TO_SQL.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class ValuesController : ControllerBase
	{
		private readonly IConfiguration configuration; //appsetting access garna use gareko interface

		public ValuesController(IConfiguration configuration)
		{
			this.configuration = configuration;
		}

		[HttpPost]
		public ActionResult import(IFormFile file) //importing file 
		{
			List<Football> list = new List<Football>();
			MemoryStream stream = new MemoryStream(); //converting file to memorystream
			file.CopyTo(stream);


			//EPPlus package use
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			ExcelPackage package = new ExcelPackage(stream);  //accessing worksheet collection 
			ExcelWorksheet worksheet = package.Workbook.Worksheets[0];//first worksheet access gareko
			var row = worksheet.Dimension.Rows; //Finding total number of rows in the worksheet

			for(int i=2; i<=row; i++)
			{
				list.Add(new Football
				{
						Id = worksheet.Cells[i, 1].GetValue<int>(), //First ma (int)worksheet.cells[i,1].value; yo garda null reference error ayo did it this way
						CustomerCode = worksheet.Cells[i, 2].GetValue<int>(),
						FirstName = worksheet.Cells[i, 3].GetValue<string>(),
						LastName = worksheet.Cells[i,4].GetValue<string>(),
						Gender = worksheet.Cells[i, 5].GetValue<string>(),
						Country = worksheet.Cells[i, 6].GetValue<string>(),
						Age = worksheet.Cells[i, 7].GetValue<int>(),
					}
					) ;
			}
			
			var connection = new SqlConnection(configuration.GetConnectionString("defaultconnection"));
			connection.ExecuteAsync("insert into customertable (CustomerCode, FirstName, LastName, Gender, Country,Age) values (@CustomerCode, @FirstName, @LastName, @Gender, @Country, @Age)", list);
			return Ok( list );
		}
	
	}
}

