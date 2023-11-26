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
using WebDriverManager.DriverConfigs.Impl;
using OpenQA.Selenium.Support.UI;


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
			finally
			{
				Console.WriteLine("\n\ndasrulebistvis daachiret sasurvel klavishs!!!");

				Console.ReadKey();
			}


		}


		public static async Task SendData(List<string> Trackings)
		{ 

			try
			{

				Console.Write("sheiyvanet wamebis raodenoba (wamebshi):");

				

				double seconds = Convert.ToDouble(Console.ReadLine());

				int miliseconds = (int)(seconds * 1000);

				Console.Write("Sheiyvanet Paroli:");

				string password = Console.ReadLine();




				using (var driver = new ChromeDriver())
				{
					driver.Navigate().GoToUrl("https://decl.rs.ge/decls.aspx");

					driver.Manage().Window.Maximize();


					var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3000));

					IWebElement usernameField = driver.FindElement(By.Id("username"));
					IWebElement passwordField = driver.FindElement(By.Id("password"));
					IWebElement loginButton = driver.FindElement(By.Id("btnLogin"));

					await Task.Delay(2000);

					usernameField.SendKeys("404640411");
					passwordField.SendKeys(password);

					//167075
					await Task.Delay(3000);

					loginButton.Click();

					await Task.Delay(6000);


					IWebElement OpenPage = driver.FindElement(By.ClassName("divModuleName"));
					await Task.Delay(5000);

					OpenPage.Click();

					await Task.Delay(5000);



					List<IWebElement> openModals = wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id("control_0_smt"))).ToList();



					await Task.Delay(3000);

					foreach (var modal in openModals)
					{
					
						string a = modal.GetAttribute("innerHTML");
						Console.WriteLine(a);
						if (modal.GetAttribute("innerHTML") == "<div>დარიდერება</div>")
						{ 
							modal.Click();
							break;
						}

					}

					await Task.Delay(30000);


					IWebElement inputTag = driver.FindElement(By.ClassName("scan_postnumber"));
					await Task.Delay(3000);

					if (inputTag == null)
					{
						Console.WriteLine("tag not found");
					}
					else
					{
						

						foreach (var tracking in Trackings)
						{
							inputTag.Clear();
							await Task.Delay(miliseconds);

							inputTag.SendKeys(tracking);
							await Task.Delay(miliseconds);

							inputTag.SendKeys(Keys.Enter);

							await Task.Delay(miliseconds);
						}
						await Task.Delay(5000);
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

