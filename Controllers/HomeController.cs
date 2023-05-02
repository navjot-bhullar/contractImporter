using ContactImport.Models;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using GemBox.Spreadsheet;
using System;
using System.Text;

namespace ContactImport.Controllers
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
        [HttpPost]
        public IActionResult Index(IFormFile postedFile)
        {
            if (postedFile != null)
            {
                //Create a Folder.
                var sb = new StringBuilder();
                string inputPath = Path.Combine("input");
                if (!Directory.Exists(inputPath))
                {
                    Directory.CreateDirectory(inputPath);
                }

                //Save the uploaded Excel file.
                string fileName = Path.GetFileName(postedFile.FileName);
                string filePath = Path.Combine(inputPath, fileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    postedFile.CopyTo(stream);
                }
                //Get the file extension  
                string fileExtension = Path.GetExtension(postedFile.FileName);
                List<ContractModel> iDataList = new List<ContractModel>();
                // If you are using the Professional version, enter your serial key below.
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                var workbook = ExcelFile.Load(filePath);
                foreach (var worksheet in workbook.Worksheets)
                {
                    sb.AppendLine();
                    sb.AppendFormat("{0} {1} {0}", new string('-', 25), worksheet.Name);

                    // Iterate through all rows in an Excel worksheet.
                    int i = 0;
                    foreach (var row in worksheet.Rows)
                    {


                        // Iterate through all allocated cells in an Excel row.
                        if (i > 0)
                        {
                            ContractModel contractObj = new ContractModel();
                            // Iterate through all allocated cells in an Excel row and update the contract object.
                            int j = 0;
                            foreach (var cell in row.AllocatedCells)
                            {
                                if (j == 0)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.Client = cell.StringValue;
                                    else
                                        contractObj.Client = "";

                                }
                                else if (j == 1)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.SingleMaster = cell.StringValue;
                                    else
                                        contractObj.SingleMaster = "";
                                }
                                else if (j == 2)
                                {


                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.JointVenture = cell.StringValue;
                                    else
                                        contractObj.JointVenture = cell.StringValue;
                                }
                                else if (j == 3)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.Name = cell.StringValue;
                                    else
                                        contractObj.Name = "";
                                }
                                else if (j == 4)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.ShortName = cell.StringValue;
                                    else
                                        contractObj.ShortName = cell.StringValue;
                                }
                                else if (j == 5)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.ContactNumber = cell.StringValue;

                                    else
                                        contractObj.ContactNumber = "";

                                }
                                else if (j == 7)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.StartDate = cell.StringValue;

                                    else
                                        contractObj.StartDate = "";

                                }
                                else if (j == 8)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.EndDate = cell.StringValue;


                                    else
                                        contractObj.EndDate = "";

                                }
                                else if (j == 9)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.ContractManager = cell.StringValue;
                                    else
                                        contractObj.ContractManager = cell.StringValue;
                                }
                                else if (j == 10)
                                {
                                    if (cell.ValueType != CellValueType.Null)
                                        contractObj.TimesheetVersionType = cell.StringValue;

                                    else
                                        contractObj.TimesheetVersionType = "";

                                }
                                j++;
                               

                            }
                            iDataList.Add(contractObj);
                        }
                        i++;
                    }
                }

                
            }


       
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
    }
}