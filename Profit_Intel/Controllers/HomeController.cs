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
        public IActionResult Portfolio()
        {
            return View();
        }

        public IActionResult Taxes()
        {
            return View();
        }

        [HttpPost]
        [ActionName("Portfolio")]
        public IActionResult Portfolio(Object input)
        {
            double portfolioVal = 0.0;
            double capitalGains = 0.0;
            ArrayList stockGains = new ArrayList();

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
            currentDirectory = Directory.GetCurrentDirectory() + "/Data/stockInfo.txt";
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
                sb.Append("<th>" + "Market Price" + "</th>");
                sb.Append("<th>" + "Number of Shares" + "</th>");
                sb.Append("<th>" + "Avg. Cost" + "</th>");
                sb.Append("<th>" + "Total Cost" + "</th>");
                sb.Append("<th>" + "Market Value" + "</th>");
                sb.Append("</tr>");
            }

            //Traverse through each stock and calculate its info 
            for (int i = 0; i < symbols.Count(); i++)
            {
                double quantity = 0.0;
                double marketVal = 0.0;
                double totalCost = 0.0;
                double gainloss = 0.0;
                double avgCost;
        
                sb.Append("<tr>");

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
                                double transactionPrice = System.Convert.ToDouble(stockInfo[k + 3]);
                                double cost = amount * transactionPrice;
                                totalCost += cost;
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
                                double transactionPrice = System.Convert.ToDouble(stockInfo[k + 3]);
                                double cost = amount * transactionPrice;
                                totalCost -= cost;
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

                //Checking if user owns any of this stock (Eliminate it from portolio if Quantity = 0)
                if (quantity > 0)
                {
                    //listing the stock symbol itself in the row
                    sb.Append("<th>" + symbols[i] + "</th>");
                    //Getting market price and adjusting portfolio and holding value accordingly
                    try
                    {
                        double price = System.Convert.ToDouble(stockPrices[(i * 2) + 1]);
                        marketVal = quantity * price;
                        portfolioVal += marketVal;
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Input string is not a sequence of digits.");
                    }
                    catch (OverflowException)
                    {
                        Console.WriteLine("The number cannot fit in a double.");
                    }

                    //Average cost basis calculation
                    avgCost = totalCost / quantity;

                    //Calculate gain or loss based on market value and total cost
                    gainloss = marketVal - totalCost;
                    //Round values to 2 decimal places
                    marketVal = Math.Round(marketVal, 2);
                    quantity = Math.Round(quantity, 2);
                    avgCost = Math.Round(avgCost, 2);
                    totalCost = Math.Round(totalCost, 2);

                    //Market price of stock
                    sb.Append("<td>" + stockPrices[(i * 2) + 1] + "</td>");
                    //Number of shares
                    sb.Append("<td>" + quantity + "</td>");
                    //Average cost
                    sb.Append("<td>" + avgCost + "</td>");
                    //Total cost of this stock
                    sb.Append("<td>" + totalCost + "</td>");
                    //Total market value of this stock
                    sb.Append("<td>" + marketVal + "</td>");
                   
                    sb.Append("</tr>");
                }

                //If the quantity of a share is 0
                else
                {
                    sb.Append("<tr>");
                    sb.Append("</tr>");
                }
            }


            sb.Append("</table>");

            portfolioVal = Math.Round(portfolioVal, 2);
            sb.Append(" <h2> Portfolio Value: $" + portfolioVal + " (USD)</h2>");

            //Profit/loss table
            sb.Append("<table class='table'><tr>");
            //Table header: Symbol and then a header for each category of info
            {
                sb.Append("<tr>");
                sb.Append("<th>" + "Symbol" + "</th>");
                sb.Append("<th>" + "Loss/Gain" + "</th>");
            }

            //Traverse through each stock and calculate its info 
            for (int i = 0; i < symbols.Count(); i++)
            {
                double quantity = 0.0;
                double marketVal = 0.0;
                double totalCost = 0.0;
                double gainloss = 0.0;
                double avgCost;

                sb.Append("<tr>");

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
                                double transactionPrice = System.Convert.ToDouble(stockInfo[k + 3]);
                                double cost = amount * transactionPrice;
                                totalCost += cost;
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
                                double transactionPrice = System.Convert.ToDouble(stockInfo[k + 3]);
                                double cost = amount * transactionPrice;
                                totalCost -= cost;
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

                //Getting market price and adjusting portfolio and holding value accordingly
                try
                {
                    double price = System.Convert.ToDouble(stockPrices[(i * 2) + 1]);
                    marketVal = quantity * price;
                    portfolioVal += marketVal;
                }
                catch (FormatException)
                {
                    Console.WriteLine("Input string is not a sequence of digits.");
                }
                catch (OverflowException)
                {
                    Console.WriteLine("The number cannot fit in a double.");
                }

                //Average cost basis calculation
                avgCost = totalCost / quantity;

                //Calculate gain or loss based on market value and total cost
                gainloss = marketVal - totalCost;
                //Round values to 2 decimal places
                marketVal = Math.Round(marketVal, 2);
                quantity = Math.Round(quantity, 2);
                avgCost = Math.Round(avgCost, 2);
                totalCost = Math.Round(totalCost, 2);
                gainloss = Math.Round(gainloss, 2);

                //Adding gains for tax data
                capitalGains += gainloss;

                //Number of shares
                //Gain or loss on this stock color based on the amount
                if (gainloss >= 0.0)
                {
                    sb.Append("<td> <font color=\"green\">+" + gainloss + "</font> </td>");
                }
                else
                {
                    sb.Append("<td> <font color=\"red\">" + gainloss + "</font> </td>");
                }
                sb.Append("</tr>");
            }

            sb.Append("</table>");

            //Saving capitl gains for tax calculation
            stockGains.Add(capitalGains.ToString());
            DataSaveWrite.WriteDataToFile(stockGains, "stockGainInfo");
            return this.Content(sb.ToString());
    }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        //Excel file upload code adapted from https://github.com/talkingdotnet/ImportExportExcelASPNetCore
        [HttpPost]
        [ActionName("ImportExport")]
        public IActionResult ImportExport(IFormFile files)
        {
            ArrayList outArr = new ArrayList();
            ArrayList stockSymbol = new ArrayList();

            IFormFile file = Request.Form.Files[0];
            if (file != null)
            {
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
                        sb.Append("<table class='table table-hover'><tr>");
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

                                    //Check if it's a unique stock Symbol (if so add to stockSymbol arrayList)
                                    if (j == 1 && !stockSymbol.Contains(row.GetCell(j).ToString()))
                                    {
                                        stockSymbol.Add(row.GetCell(j).ToString());
                                    }
                                }
                            }
                            sb.AppendLine("</tr>");
                        }
                        sb.Append("</table>");
                        stream.Close();             // Close the reading stream
                    }

                }
                Debug.WriteLine(outArr.ToArray());
                DataSaveWrite.WriteDataToFile(outArr, "stockInfo");     // Write it to a file
                DataSaveWrite.WriteDataToFile(stockSymbol, "stockList");
                return this.Content(sb.ToString());
            }

            return new EmptyResult();
        }

        [HttpGet]
        public IActionResult ImportExport()
        { return View(); }

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


        // In the future this would be made so that users can give us the name of the specific file in which 
        // they would have their stockList input, like some form of ID, now it just grabs the generic one
        [HttpGet]
        public IActionResult stockList()
        {
            StringBuilder sb = new StringBuilder();
            StreamReader reader = System.IO.File.OpenText("Data\\stockList.txt");
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                sb.Append(line+"\n");
            }
            reader.Close();
            return this.Content(sb.ToString());
        }
    } }
