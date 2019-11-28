using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using Profit_Intel.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Html;
using System.Collections;
using DataAnalysis;

namespace Profit_Intel.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Analyze()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }
        [HttpGet]
        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }


        /*
         *  This is where Eddys change was
         */
        [HttpPost]
        [ActionName("Contact")]
        public IActionResult Contact(Object input)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Hello World!");
            return this.Content(sb.ToString());
        }
        public IActionResult Portfolio()
        {
            return View();
        }

        [HttpPost]
        [ActionName("Portfolio")]
        public IActionResult Portfolio(Object input)
        {
            double portfolioVal = 0.0;

            StringBuilder sb = new StringBuilder();
            sb.Append("<table class='table'><tr>");

            //Directory containing saved stock file
            string currentDirectory = Directory.GetCurrentDirectory() + "/Data/stockList.txt";
            //Creating an array in which we store the read conent of text file containing stock symbols
            string[] symbols;
            var list = new List<string>();
            var fileStream = new FileStream(currentDirectory, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            symbols = list.ToArray();

            //Directory containing saved stock info
            currentDirectory = Directory.GetCurrentDirectory() + "/Data/StockInfo.txt";
            //Creating an array in which we store the read conent of text file containing all stock info
            string[] stockInfo;
            list = new List<string>();
            fileStream = new FileStream(currentDirectory, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            stockInfo = list.ToArray();

            //Directory containing saved stock prices
            currentDirectory = Directory.GetCurrentDirectory() + "/Data/stockPrices.txt";
            //Creating an array in which we store the read conent of text file containing all stock info
            string[] stockPrices;
            list = new List<string>();
            fileStream = new FileStream(currentDirectory, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            stockPrices = list.ToArray();

            //Table header: Symbol and then a header for each category of info
            {
                sb.Append("<tr>");
                sb.Append("<th>" + "Symbol" + "</th>");
                sb.Append("<th>" + "Quantity" + "</th>");
                sb.Append("<th>" + "Market Price" + "</th>");
                sb.Append("<th>" + "Holding Value" + "</th>");
                sb.Append("</tr>");
            }

            //Traverse through each stock and calculate its info 
            for (int i = 0; i < symbols.Count(); i++)
            {
                double quantity = 0.0;
                double holdingVal = 0.0;

                sb.Append("<tr>");

                //Every entry in the array thats has remainder 1 when mod 5 is a symbol
                //  if(i%5 == 1 && !stockNames.Contains(lines[i]))

                //listing the stock symbol itself in the row
                sb.Append("<th>" + symbols[i] + "</th>");

                //Calculating quantity of this stock by iterating through stockInfo
                for (int k = 0; k < stockInfo.Count(); k++)
                {
                    //Checking if an entry matches this stock symbol
                    if (stockInfo[k].Equals(symbols[i]))
                    {
                        //If an entry is a BUY
                        if (stockInfo[k + 1].Equals("BUY"))
                        {
                            try
                            {
                                double amount = System.Convert.ToDouble(stockInfo[k + 2]);
                                quantity += amount;
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine("Input string is not a sequence of digits.");
                            }
                            catch (OverflowException)
                            {
                                Console.WriteLine("The number cannot fit in a double.");
                            }
                        }

                        //If an entry is a SELL
                        if (stockInfo[k + 1].Equals("SELL"))
                        {
                            try
                            {
                                double amount = System.Convert.ToDouble(stockInfo[k + 2]);
                                quantity -= amount;
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine("Input string is not a sequence of digits.");
                            }
                            catch (OverflowException)
                            {
                                Console.WriteLine("The number cannot fit in a double.");
                            }
                        }
                    }
                }

                quantity = Math.Round(quantity, 2);
                sb.Append("<td>" + quantity + "</td>");
                sb.Append("<td>" + stockPrices[(i * 2) + 1] + "</td>");

                try
                {
                    double price = System.Convert.ToDouble(stockPrices[(i * 2) + 1]);
                    holdingVal = quantity * price;
                    holdingVal = Math.Round(holdingVal, 2);
                    portfolioVal += holdingVal;
                }
                catch (FormatException)
                {
                    Console.WriteLine("Input string is not a sequence of digits.");
                }
                catch (OverflowException)
                {
                    Console.WriteLine("The number cannot fit in a double.");
                }

                sb.Append("<td>" + holdingVal + "</td>");

                sb.Append("</tr>");
            }


            sb.Append("</table>");

            portfolioVal = Math.Round(portfolioVal, 2);
            sb.Append(" <h2> Portfolio Value: $" + portfolioVal + " (USD)</h2>");
            return this.Content(sb.ToString());
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        [HttpPost]
        [ActionName("ImportExport")]
        public IActionResult ImportExport(IFormFile files)
        {
            ArrayList outArr = new ArrayList();
            IFormFile file = Request.Form.Files[0];
            if (file != null) {
                string folderName = "Upload";
                string webRootPath = _hostingEnvironment.WebRootPath;
                string newPath = Path.Combine(webRootPath, folderName);
                StringBuilder sb = new StringBuilder();
                if (!Directory.Exists(newPath))
                {
                    Directory.CreateDirectory(newPath);
                }
                if (file.Length > 0)
                {
                    string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                    ISheet sheet;
                    string fullPath = Path.Combine(newPath, file.FileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                        stream.Position = 0;
                        if (sFileExtension == ".xls")
                        {
                            HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                            sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  
                        }
                        else
                        {
                            XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                            sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook   
                        }

                        IRow headerRow = sheet.GetRow(0); //Get Header Row
                        int cellCount = headerRow.LastCellNum;
                        sb.Append("<table class='table'><tr>");
                        for (int j = 0; j < cellCount; j++)
                        {
                            NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                            if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                            outArr.Add(cell.ToString());                    // Add to the array
                            sb.Append("<th>" + cell.ToString() + "</th>");
                        }
                        sb.Append("</tr>");
                        sb.AppendLine("<tr>");
                        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue;
                            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    outArr.Add(row.GetCell(j).ToString());
                                    sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");
                                }
                            }
                            sb.AppendLine("</tr>");
                        }
                        sb.Append("</table>");
                    }
                }
                Debug.WriteLine(outArr.ToArray());
                DataSaveWrite.WriteDataToFile(outArr, "StockInfo");     // Write it to a file
                return this.Content(sb.ToString());
            } 

            return new EmptyResult();
        }

        [HttpGet]
        public IActionResult ImportExport()
        { return View(); }
    
    } }
