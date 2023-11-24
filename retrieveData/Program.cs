using System;
using System.Numerics;
using OfficeOpenXml;
using System.Net.Http;
using System.Collections.Generic;
using System.Net;
using HtmlAgilityPack;

//C:\Users\JAJO\Downloads tracking
namespace retrievedata
{
	class dataClass
	{
		public static async Task Main()
		{

			/*Console.Write("Input your Path: ");*/

			/*string pathtofile = Console.ReadLine();

			Console.Write("File name:");

			string filename = Console.ReadLine();*/

			string path = @"C:\Users\JAJO\Downloads\tracking.xlsx";


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
				
					await SendData(Trackings);

				}



			}
			catch (Exception ex)
			{

				Console.WriteLine($"An error occurred: {ex.Message}");
			}
		}

		public static async Task SendData(List<string> Trackings)
		{
			
			try
			{
				using (var httpClient = new HttpClient())
				{
					var Url = "http://localhost:5500/index";

					var Data = new FormUrlEncodedContent(new Dictionary<string, string>
			{
				{"Trackings", string.Join(".", Trackings) }
			});


					var response = await httpClient.PostAsync(Url, Data);

					var responseBody = await response.Content.ReadAsStringAsync();

				

					
					if (response.IsSuccessStatusCode)
					{
						Console.WriteLine($"Response from website: {responseBody}");
						Console.WriteLine("Data Retrieved Successfully");

						HttpResponseMessage responseContent = await httpClient.GetAsync(Url);

						if (responseContent.IsSuccessStatusCode)
						{
							var htmlContent = await responseContent.Content.ReadAsStringAsync();
							Console.WriteLine("HTML Content:");
							Console.WriteLine(htmlContent);

							var htmldom = new HtmlDocument();
							htmldom.LoadHtml(htmlContent);

							var inputTag = htmldom.DocumentNode.SelectSingleNode("//*[@id = 'trId']");

							if(inputTag != null)
							{
								foreach(var tracking in Trackings)
								{
									Console.WriteLine(tracking);
									inputTag.SetAttributeValue("value", tracking);

									System.Threading.Thread.Sleep(1000);
									inputTag.SetAttributeValue("value", null);
									System.Threading.Thread.Sleep(1000);
								}
							}
							else
							{
								Console.WriteLine("tag not found");
							}
						
						}

					}	
					else
					{
						Console.WriteLine($"Error: Status Code {response.StatusCode}");
						Console.WriteLine($"Response Body: {responseBody}");
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"{ex.Message}   Error");

			}
		}
		


	}


}