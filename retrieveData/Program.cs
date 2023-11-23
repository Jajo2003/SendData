using System;
using System.Numerics;
using OfficeOpenXml;
using System.Net.Http;
using System.Collections.Generic;
namespace retrievedata
{
	class dataClass
	{
		public static async Task Main()
		{

			Console.Write("Input your Path: ");

			string pathtofile = Console.ReadLine();

			Console.Write("File name:");

			string filename = Console.ReadLine();

			string path = pathtofile + @"\" +filename +".xlsx";


			List<string> Trackings = new List<string>();

			try
			{
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
				using (var package = new ExcelPackage(new FileInfo(path)))
				{

					var sheet = package.Workbook.Worksheets[0];


					int rowCount = sheet.Dimension.Rows;

					//Cikli amatebs Excel failshi arsebul informacias.
					for (int i = 1; i <= rowCount; i++)
					{
						Trackings.Add(sheet.Cells["A" + $"{i}"].Text);
					}
					for (int i = 1; i <= rowCount; i++)
					{
						Console.WriteLine(Trackings[i]);
					}



					await SendData(Trackings);

				}



			}
			catch (Exception ex)
			{

				Console.WriteLine($"An error occurred: {ex.Message}");
			}
		}

		private static async Task SendData(List<string> Trackings)
		{

			try
			{
				using (var httpClient = new HttpClient())
				{
					var Url = "http://localhost:5500/AJAX/";

					var Data = new FormUrlEncodedContent(new Dictionary<string, string>{

						{"Trackings",string.Join(".",Trackings) }

					});
					var response = await httpClient.PostAsync(Url, Data);

					response.EnsureSuccessStatusCode();

					var responseBody = await response.Content.ReadAsStringAsync();
					Console.WriteLine($"Response from website : ${responseBody}");
					Console.WriteLine("Data Retrieved Succesfully");

				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"{ex.Message}   Error");
			}

		}







	}


}