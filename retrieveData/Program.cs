using System;
using System.Numerics;
using OfficeOpenXml;
using System.Net.Http;
using System.Collections.Generic;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;

namespace retrievedata
{
	class dataClass
	{
		public static async Task Main()
		{

			Console.Write("Input your Path:");

			string pathtofile = Console.ReadLine();

			Console.Write("File name:");

			string filename = Console.ReadLine();

			string path = pathtofile +@"\"+ filename +".xlsx";

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


				using (var driver = new ChromeDriver())
				{
					driver.Navigate().GoToUrl("http://localhost:5500/index");

					IWebElement inputTag = driver.FindElement(By.Id("trId"));

					if (inputTag == null)
					{
						Console.WriteLine("tag not found");
					}
					else
					{
						foreach (var tracking in Trackings)
						{
							Console.WriteLine(tracking);

							inputTag.SendKeys(tracking);
							await Task.Delay(700);

							inputTag.Clear();
							await Task.Delay(700);
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"{ex.Message} Error");
			}
		}






	}



	/*

	C:\Users\JAJO\Downloads
	tracking


	 */


}

