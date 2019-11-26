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
