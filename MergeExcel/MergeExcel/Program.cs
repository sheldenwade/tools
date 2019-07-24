using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Drawing;

namespace MergeExcel
{
    class PrimaryGroup
    {
        public string GroupName { get; set; }
        public int Total { get; set; }
        public int Passed { get; set; }
        public double PassedRate { get; set; }
        public bool IsUsed { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string startCellStringTemplateOppo = @"B{0}";
            string endCellStringTemplateOppo = @"H{0}";

            string startCellStringTemplatePenetration = @"B14";
            string endCellStringTemplatePenetration = @"D19";

            string baseFolder = ConfigurationManager.AppSettings["BaseFolder"];
            string inputFolder = ConfigurationManager.AppSettings["InputFolder"];

            var excelConfigs = File.ReadAllLines("MappingFile.txt")
                .Where(i => !i.StartsWith("#") && !string.IsNullOrEmpty(i))
                .Select(i => i.Split('\t'))
                .Select(i => new
                {
                    ExcelName = i[0],
                    FileType = i[1],
                    WorkSheetName = i[2]
                })
                .ToDictionary(i => i.ExcelName, j => j);

            string excelTemplateFile = Path.Combine(baseFolder, "ExcelTemplate.xlsx");
            DateTime time = DateTime.Now;
            string savedExcelOutput = string.Format(@"{0}\ExcelTemplate_Result_{1:yyyyMMddTHHmmss}.xlsx", baseFolder, time);
            string errorFileOutput = string.Format(@"{0}\ErrorFile_{1:yyyyMMddTHHmmss}.txt", baseFolder, time);
            StreamWriter swErrorFile = new StreamWriter(errorFileOutput);
            Application excelApp = new Application();
            //excelApp.Visible = true;
            Workbook excelTemplateWorkbook = excelApp.Workbooks.Open(excelTemplateFile);
            Console.WriteLine("Merging Data starts...");
            foreach (var configItem in excelConfigs)
            {
                string fileName = Path.Combine(inputFolder, configItem.Key + ".xlsx");

                var config = configItem.Value;
                Console.WriteLine("Excel Name: {0} , File Type: {1} , Worksheet Name : {2}", config.ExcelName, config.FileType, config.WorkSheetName);
                if (!File.Exists(fileName))
                {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine("\t{0} Not Found!", fileName);
                    Console.ResetColor();
                    swErrorFile.WriteLine("{0} Not Found!", fileName);
                    continue;
                }

                Console.WriteLine("\t{0} Found", fileName);
                Console.WriteLine("\tProcessing " + fileName);

                Workbook excelWorkBook = excelApp.Workbooks.Open(fileName);
                string workSheetName = config.WorkSheetName;
                Worksheet excelWorkSheet = (Worksheet)excelWorkBook.Worksheets[1];
                Worksheet excelTemplateWorkSheet = (Worksheet)excelTemplateWorkbook.Worksheets[workSheetName];
                Console.WriteLine("\tCopying Data Starts");

                if (config.FileType == "Penetration")
                {
                    Range rangeFrom = excelWorkSheet.Range[startCellStringTemplatePenetration, endCellStringTemplatePenetration];
                    if (rangeFrom == null || rangeFrom.Cells[0, 0] == null || rangeFrom.Cells[0, 0].Value2 == null)
                    {
                        Console.Write("\t");
                        Console.BackgroundColor = ConsoleColor.Red;
                        Console.WriteLine("Malformed data in file: {0}", fileName);
                        swErrorFile.WriteLine("Malformed data in file: {0}", fileName);
                        Console.ResetColor();
                        excelWorkBook.Close(false, null, null);
                        Marshal.ReleaseComObject(excelWorkSheet);
                        Marshal.ReleaseComObject(excelWorkBook);
                        continue;
                    }
                    rangeFrom.Copy(Type.Missing);
                    Range rangeTo = excelTemplateWorkSheet.Range["E17"];
                    rangeTo.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
                else if (config.FileType == "Oppo")
                {
                    Console.WriteLine("\tProcessing Oppo file: {0}", fileName);
                    Console.WriteLine("\tPenetration passed rate check");
                    if (double.TryParse(GetCellValue(excelTemplateWorkSheet.Cells, 22, 7), out double penPassedRate))
                    {
                        Console.WriteLine("\tPnetration Passed Rate is {0}", penPassedRate);
                        if (penPassedRate > 80)
                        {
                            Console.BackgroundColor = ConsoleColor.Green;
                            Console.WriteLine("\tPnetration Passed Rate is bigger than 80. Skip Oppo data importation.");
                            Console.ResetColor();
                            excelWorkBook.Close(false, null, null);
                            Marshal.ReleaseComObject(excelWorkSheet);
                            Marshal.ReleaseComObject(excelWorkBook);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine("\tContinue Oppo data import");
                        }

                        int rowIndexSourceStart = 1;
                        while (excelWorkSheet.Cells[rowIndexSourceStart, 2] == null || string.IsNullOrWhiteSpace(excelWorkSheet.Cells[rowIndexSourceStart, 2].Value2))
                        {
                            rowIndexSourceStart++;
                        }

                        if (excelWorkSheet.Cells[rowIndexSourceStart, 2] == null || excelWorkSheet.Cells[rowIndexSourceStart, 2].Value2 == null)
                        {
                            Console.BackgroundColor = ConsoleColor.Red;
                            Console.WriteLine("\tMalformed data in file: {0}", fileName);
                            swErrorFile.WriteLine("Malformed data in file: {0}", fileName);
                            Console.ResetColor();
                            excelWorkBook.Close(false, null, null);
                            Marshal.ReleaseComObject(excelWorkSheet);
                            Marshal.ReleaseComObject(excelWorkBook);
                            continue;
                        }

                        int rowIndexSourceEnd = rowIndexSourceStart;
                        string checkValueSource = excelWorkSheet.Cells[rowIndexSourceEnd, 2].Value2.ToString();
                        while (!string.IsNullOrEmpty(checkValueSource))
                        {
                            rowIndexSourceEnd++;
                            checkValueSource = excelWorkSheet.Cells[rowIndexSourceEnd, 2].Value2.ToString();
                        }

                        int rowIndexDestination = 1;
                        string checkValueDestination = GetCellValue(excelTemplateWorkSheet.Cells, rowIndexDestination, 1);
                        while (checkValueDestination.ToLower() != "opportunity list" && rowIndexDestination < 3000)
                        {
                            rowIndexDestination++;
                            checkValueDestination = GetCellValue(excelTemplateWorkSheet.Cells, rowIndexDestination, 1);
                        }
                        if (checkValueDestination.ToLower() != "opportunity list")
                        {
                            Console.Write("\t");
                            Console.BackgroundColor = ConsoleColor.Red;
                            Console.WriteLine("Malformed data in file: {0}", fileName);
                            swErrorFile.WriteLine("Malformed data in file: {0}", fileName);
                            Console.ResetColor();
                            excelWorkBook.Close(false, null, null);
                            Marshal.ReleaseComObject(excelWorkSheet);
                            Marshal.ReleaseComObject(excelWorkBook);
                            continue;
                        }
                        rowIndexDestination += 2;
                        Range rangeFrom = excelWorkSheet.Range[string.Format(startCellStringTemplateOppo, rowIndexSourceStart + 1), string.Format(endCellStringTemplateOppo, rowIndexSourceEnd)];
                        Range rangeTo = excelTemplateWorkSheet.Range["A" + rowIndexDestination];
                        rangeFrom.Copy(rangeTo);
                    }
                    else
                    {
                        Console.BackgroundColor = ConsoleColor.Red;
                        Console.WriteLine("\tPnetration Rate Check failed! Please check the mapping file sequence and data source!");
                        swErrorFile.WriteLine("Pnetration Rate Check failed when processing: {0}", fileName);
                        Console.ResetColor();
                        excelWorkBook.Close(false, null, null);
                        Marshal.ReleaseComObject(excelWorkSheet);
                        Marshal.ReleaseComObject(excelWorkBook);
                        continue;
                    }
                }
                else if (config.FileType == "Listed Banner")
                {
                    Dictionary<string, PrimaryGroup> groupMap = new Dictionary<string, PrimaryGroup>();
                    int rowIndex = 14;
                    int columnIndex = 1;
                    if (excelWorkSheet.Cells == null 
                        || excelWorkSheet.Cells[rowIndex, columnIndex] == null 
                        || string.IsNullOrEmpty(excelWorkSheet.Cells[rowIndex, columnIndex].Value2))
                    {
                        Console.Write("\t");
                        Console.BackgroundColor = ConsoleColor.Red;
                        Console.WriteLine("Malformed data in file: {0}", fileName);
                        swErrorFile.WriteLine("Malformed data in file: {0}", fileName);
                        Console.ResetColor();
                        excelWorkBook.Close(false, null, null);
                        Marshal.ReleaseComObject(excelWorkSheet);
                        Marshal.ReleaseComObject(excelWorkBook);
                        continue;
                    }
                    while (excelWorkSheet.Cells != null 
                        && excelWorkSheet.Cells[rowIndex, columnIndex] != null 
                        && !string.IsNullOrEmpty(excelWorkSheet.Cells[rowIndex, columnIndex].Value2.ToString()))
                    {
                        string groupName = excelWorkSheet.Cells[rowIndex, columnIndex].Value2;
                        if (groupName.ToLower() == "totals")
                        {
                            break;
                        }
                        int total = int.Parse(excelWorkSheet.Cells[rowIndex, columnIndex + 1].Value2.ToString());
                        int passed = int.Parse(excelWorkSheet.Cells[rowIndex, columnIndex + 2].Value2.ToString());
                        double passedRate = double.Parse(excelWorkSheet.Cells[rowIndex, columnIndex + 3].Value2.ToString());
                        PrimaryGroup group = new PrimaryGroup
                        {
                            GroupName = groupName,
                            Total = total,
                            Passed = passed,
                            PassedRate = passedRate,
                            IsUsed = false
                        };
                        groupMap.Add(groupName, group);
                        rowIndex++;
                    }
                    int rowIndexTemplateListPen = 29;
                    List<PrimaryGroup> usedGroup = new List<PrimaryGroup>();
                    while (excelTemplateWorkSheet.Cells != null 
                        && excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 2] != null 
                        && !string.IsNullOrEmpty(excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 2].Value2))
                    {
                        string templateGroupName = excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 2].Value2;
                        if (groupMap.ContainsKey(templateGroupName))
                        {
                            var group = groupMap[templateGroupName];
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 5].Value2 = group.Total;
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 6].Value2 = group.Passed;
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 7].Value2 = group.PassedRate;
                            group.IsUsed = true;
                            usedGroup.Add(group);
                        }
                        rowIndexTemplateListPen++;
                    }
                    if (usedGroup.Count > 0)
                    {
                        int totalSum = usedGroup.Sum(i => i.Total);
                        int passedSum = usedGroup.Sum(i => i.Passed);
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 5].Value2 = totalSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 6].Value2 = passedSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 7].Value2 = passedSum * 100.0 / totalSum;
                    }
                    if (groupMap.Any(i => !i.Value.IsUsed))
                    {
                        int normalListPenRowIndex = rowIndexTemplateListPen;
                        rowIndexTemplateListPen++;
                        var unusedGroups = groupMap.Where(i => !i.Value.IsUsed);
                        foreach (var group in unusedGroups)
                        {
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 2].Value2 = group.Key;
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 5].Value2 = group.Value.Total;
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 6].Value2 = group.Value.Passed;
                            excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 7].Value2 = group.Value.PassedRate;
                            excelTemplateWorkSheet.Rows[rowIndexTemplateListPen + 1].Insert();
                            rowIndexTemplateListPen++;
                        }
                        int missedTotalSum = unusedGroups.Sum(i => i.Value.Total);
                        int missedPassedSum = unusedGroups.Sum(i => i.Value.Passed);
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 1].Value2 = "Missing Total";
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 5].Value2 = missedTotalSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 6].Value2 = missedPassedSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 7].Value2 = missedPassedSum * 100.0 / missedTotalSum;
                        excelTemplateWorkSheet.Rows[rowIndexTemplateListPen + 1].Insert();
                        rowIndexTemplateListPen++;

                        int totalSum = groupMap.Sum(i => i.Value.Total);
                        int passedSum = groupMap.Sum(i => i.Value.Passed);
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 1].Value2 = "Total";
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 5].Value2 = totalSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 6].Value2 = passedSum;
                        excelTemplateWorkSheet.Cells[rowIndexTemplateListPen, 7].Value2 = passedSum * 100.0 / totalSum;
                        //excelTemplateWorkSheet.Rows[rowIndexTemplateListPen + 1].Insert();
                        //rowIndexTemplateListPen++;

                        Range formatSourceRange = excelTemplateWorkSheet.Range["A" + normalListPenRowIndex, "D" + normalListPenRowIndex];
                        formatSourceRange.Copy();

                        Range formatDestRange = excelTemplateWorkSheet.Range["A" + (rowIndexTemplateListPen - 1), "D" + rowIndexTemplateListPen];
                        formatDestRange.PasteSpecial(XlPasteType.xlPasteFormats);

                        Range missedDataRange = excelTemplateWorkSheet.Range["E" + (normalListPenRowIndex + 1), "G" + rowIndexTemplateListPen];
                        missedDataRange.BorderAround(Type.Missing, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, Type.Missing);

                        Range missedDataSumRange = excelTemplateWorkSheet.Range["E" + (rowIndexTemplateListPen - 1), "G" + rowIndexTemplateListPen];
                        missedDataSumRange.Font.Bold = true;
                        missedDataSumRange.Interior.Color = ColorTranslator.ToOle(Color.WhiteSmoke);
                    }
                }

                excelWorkBook.Close(false, null, null);
                Marshal.ReleaseComObject(excelWorkSheet);
                Marshal.ReleaseComObject(excelWorkBook);
                Console.WriteLine("\tCopying Data Ends");
            }

            Console.WriteLine("Outputting the merged Excel File to: \t {0}", savedExcelOutput);
            excelTemplateWorkbook.SaveAs(savedExcelOutput,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                XlSaveAsAccessMode.xlNoChange,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
            swErrorFile.Close();
            Console.WriteLine("Merge data work done! ");
            Console.WriteLine("Press any key to close the window...");
            Console.ReadKey();
        }

        private static string GetCellValue(Range range, int rowIndex, int colIndex)
        {
            if (range == null)
            {
                return string.Empty;
            }

            if (range[rowIndex, colIndex] == null)
            {
                return string.Empty;
            }

            if (range[rowIndex, colIndex].Value2 == null)
            {
                return string.Empty;
            }

            return range[rowIndex, colIndex].Value2.ToString();
        }
    }
}
