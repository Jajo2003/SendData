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

				Console.Write("sheiyvanet wamebis raodenoba (wamebshi):");
				double seconds = Convert.ToDouble(Console.ReadLine());

				int miliseconds = (int)(seconds * 1000);





				using (var driver = new ChromeDriver())
				{
					driver.Navigate().GoToUrl("https://decl.rs.ge/decls.aspx");

					IWebElement inputTag = driver.FindElement(By.Id("decl_input_t scan_postnumber"));

					if (inputTag == null)
					{
						Console.WriteLine("tag not found");
					}
					else
					{
						

						foreach (var tracking in Trackings)
						{ 

							inputTag.SendKeys(tracking);
							await Task.Delay(miliseconds);

							inputTag.Clear();
							await Task.Delay(miliseconds);
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






}

