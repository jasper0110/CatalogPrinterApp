using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtil
{
    public enum CatalogType
    {
        PARTICULIER = 1,
        DAKWERKER = 2,
        VERANDA = 3,
        AANNEMER = 4,
        BLANCO = 5,
        STOCK = 6
    }

    public struct InputRanges
    {       
        public string catalogType;
        public string btw;
        public string footerRight;
        public string footerLeft;
        public string footerMidFirst;
        public string footerMidSecond;
        public string headerRight;
        public string headerLeft;
        public string headerMid;
        public string printArea;
    }

    public static class ExcelUtility
    {
        public static readonly string _tmpWorkbookDir = @"C:\temp";
        public static readonly string _tmpWokbookName = @"temp.xlsx";

        public static Workbook MasterWb { get; set; }
        public static Workbook Wb2Print { get; set; }

        public static Application XlApp { get
            {
                try
                {
                    var excel = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    excel.DisplayAlerts = false;
                    excel.Visible = false;
                    return excel;
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("0x800401E3 (MK_E_UNAVAILABLE)"))
                    {
                        var excel = new Application();
                        excel.DisplayAlerts = false;
                        excel.Visible = false;
                        return excel;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        private static List<string> Range2List(string range)
        {
            var list = new List<string>();

            var first = range.Substring(0, range.IndexOf("-"));
            var second = range.Substring(range.IndexOf("-") + 1);

            var firstInt = 0;
            var secondInt = 0;
            if (!Int32.TryParse(first, out firstInt))
                throw new Exception($"Parse exception for sheet input " + first + " from range " + range + "! Please check sheet input.");
            if (!Int32.TryParse(second, out secondInt))
                throw new Exception($"Parse exception for sheet input " + second + " from range " + range + "! Please check sheet input.");

            for (int i = firstInt; i <= secondInt; ++i)
                list.Add(i.ToString());

            return list;
        }

        private static List<string> MultipleRange2List(string sheetInput)
        {
            var sheets2Print = new List<string>();

            var sheets = sheetInput.Split(';').ToList();
            foreach (var sheet in sheets)
            {
                if (sheet.Contains("-"))
                {
                    sheets2Print.AddRange(Range2List(sheet));
                }
                else
                {
                    if (sheet.Length > 0)
                        sheets2Print.Add(sheet);
                }
            }
            return sheets2Print;
        }

        private static List<string> GetSheetsFromInputPages(string sheetInput, string firstPage)
        {
            var masterWbSheetOrder = new List<string>();

            Worksheet sh = MasterWb.Sheets[firstPage];
            while(sh != null)
            {
                masterWbSheetOrder.Add(sh.Name);
                sh = sh.Next;
            }

            var inputPages = MultipleRange2List(sheetInput).Select(s => Int32.Parse(s));
            int maxIndex = inputPages.Max();
            if(maxIndex > masterWbSheetOrder.Count)
                throw new Exception($"Page input " + maxIndex + " not found in " + MasterWb + "! Please check page input.");

            var sheets2Print = new List<string>();
            foreach(var input in inputPages)
            {
                sheets2Print.Add(masterWbSheetOrder[input - 1]);
            }

            return sheets2Print;
        }

        public static KeyValuePair<int, int> StringRange2Coordinate(string range)
        {
            var column = new string(range.TakeWhile(char.IsUpper).ToArray());
            if (column.Length < 1)
                return new KeyValuePair<int, int>(0,0);

            var row = range.Substring(column.Length);

            int columnInt = 0;
            foreach(char c in column)
            {
                columnInt += char.ToUpper(c) - 64;
            }

            int rowInt = 0;
            if(!Int32.TryParse(row, out rowInt))
                throw new Exception($"Parse exception for row " + row + " from range " + range + "!");

            return new KeyValuePair<int, int>(rowInt, columnInt);
        }

        public static void ChangePassword(string wbFullName, string oldPassword, string newPassword)
        {
            Workbook wb = ExcelUtility.GetWorkbook(wbFullName, oldPassword);
            if (wb == null)
                throw new Exception($"Wrong password for workbook " + wbFullName + "!");
            wb.Password = newPassword;
            ExcelUtility.CloseWorkbook(wb, true);
        }

        public static Workbook GetWorkbook(string fullName, string password = null)
        {
            try
            {
                if (password != null)
                    return XlApp.Workbooks.Open(fullName, Password: password);

                return XlApp.Workbooks.Open(fullName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        public static Worksheet GetWorksheetByName(Workbook wb, string shName)
        {
            return wb.Worksheets.OfType<Worksheet>().FirstOrDefault(ws => ws.Name == shName);
        }

        public static void CloseWorkbook(Workbook wb, bool saveChanges)
        {
            if (wb != null)
            {
                wb.Close(saveChanges);
                Marshal.ReleaseComObject(wb);
            }
        }

        public static void CloseExcel()
        {
            XlApp.Quit();
            Marshal.ReleaseComObject(XlApp);
        }

        public static bool IsFileInUse(string path)
        {
            FileStream stream = null;

            try
            {
                stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (System.IO.FileNotFoundException)
            {
                // file does not exist
                return false;
            }
            catch (Exception e)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread 
                throw new Exception("FileInUseException!" , e);
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

        public static void ExportWorkbook2Pdf(Progress<int> progress, string wbName, string password, string catalogType, string outputPath, string sheetInput, string firstPage, InputRanges ranges, bool inclBtw)
        {
            try
            { 
                // open master workbook
                MasterWb = ExcelUtility.GetWorkbook(wbName, password);

                // open temp workbook to which the sheets of interest are copied to
                Wb2Print = ExcelUtility.XlApp.Workbooks.Add();
                if (!Directory.Exists(_tmpWorkbookDir))
                    Directory.CreateDirectory(_tmpWorkbookDir);
                Wb2Print?.SaveAs(_tmpWorkbookDir + @"\" + _tmpWokbookName);

                // progress update
                ((IProgress<int>)progress).Report(30);

                string leftHeader = "", centerHeader = "", rightHeader = "", leftFooter = "", centerFooterFirst = "", centerFooterSecond = "", rightFooter = "";

                var cellFooterRight = StringRange2Coordinate(ranges.footerRight);
                if (cellFooterRight.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterRight: " + ranges.footerRight + "!");
                var cellFooterLeft = StringRange2Coordinate(ranges.footerLeft);
                if (cellFooterLeft.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterLeft: " + ranges.footerLeft + "!");
                var cellFooterMidFirst = StringRange2Coordinate(ranges.footerMidFirst);
                if (cellFooterMidFirst.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterMidFirst: " + ranges.footerMidFirst + "!");
                var cellFooterMidSecond = StringRange2Coordinate(ranges.footerMidSecond);
                if (cellFooterMidSecond.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterMidSecond: " + ranges.footerMidSecond + "!");
                var cellHeaderRight = StringRange2Coordinate(ranges.headerRight);
                if (cellHeaderRight.Key == 0)
                    throw new Exception($"Invalid input for cell range cellHeaderRight: " + ranges.headerRight + "!");
                var cellHeaderLeft = StringRange2Coordinate(ranges.headerLeft);
                if (cellHeaderLeft.Key == 0)
                    throw new Exception($"Invalid input for cell range cellHeaderLeft: " + ranges.headerLeft + "!");
                var cellHeaderMid = StringRange2Coordinate(ranges.headerMid);
                if (cellHeaderMid.Key == 0)
                    throw new Exception($"Invalid input for cell range cellHeaderMid: " + ranges.headerMid + "!");

                var cellCatalogType = StringRange2Coordinate(ranges.catalogType);
                var cellBtw = StringRange2Coordinate(ranges.btw);

                // get sheet order
                var sheetOrder = MultipleRange2List(sheetInput);

                // catalog int type
                int catalogTypeInt;
                switch (catalogType)
                {
                    case "Particulier":
                        catalogTypeInt = (int)CatalogType.PARTICULIER;
                        break;
                    case "Dakwerker":
                        catalogTypeInt = (int)CatalogType.DAKWERKER;
                        break;
                    case "Veranda":
                        catalogTypeInt = (int)CatalogType.VERANDA;
                        break;
                    case "Aannemer":
                        catalogTypeInt = (int)CatalogType.AANNEMER;
                        break;
                    case "Blanco":
                        catalogTypeInt = (int)CatalogType.BLANCO;
                        break;
                    case "Stock":
                        catalogTypeInt = (int)CatalogType.STOCK;
                        break;
                    default:
                        catalogTypeInt = 0;
                        break;
                   
                }

                double progressSheets = 30.0;
                double incr = 70.0 / (double)sheetOrder.Count;

                // copy necessary sheets to temp workbook and put sheets in correct order
                foreach (var shName in sheetOrder)
                {
                    if (ExcelUtility.GetWorksheetByName(MasterWb, shName) == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + MasterWb + "!" +
                            "\nPlease check the sheet order input.");

                    // set catalog type
                    MasterWb.Sheets[shName].Cells[cellCatalogType.Key, cellCatalogType.Value] = catalogTypeInt;
                    // set btw
                    if (inclBtw)
                    {
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 1;
                    }
                    else
                    {
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 2;
                    }

                    // get headers and footers
                    leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "";
                    centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "";
                    rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "";
                    leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "";
                    centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "";
                    centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "";
                    rightFooter = "TARIEF Nr. " + shName;

                    // unhide all columns
                    MasterWb.Sheets[shName].Cells(1, 1).EntireRow.EntireColumn.Hidden = false;

                    // copy sheet
                    MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                    //format sheet
                    FormatSheet(Wb2Print.Sheets[shName], leftHeader, centerHeader, rightHeader, leftFooter, centerFooterFirst, centerFooterSecond, rightFooter, ranges.printArea);

                    //// copy sheet
                    //if (catalogTypeInt == (int)CatalogType.PARTICULIER)
                    //{
                    //    // set btw false
                    //    MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 2;

                    //    // get headers and footers
                    //    leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "";
                    //    centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "";
                    //    rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "";
                    //    leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "";
                    //    centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "";
                    //    centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "";
                    //    rightFooter = "TARIEF Nr. " + shName;

                    //    // copy sheet
                    //    MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                    //    //format sheet
                    //    FormatSheet(Wb2Print.Sheets[shName + " (2)"], leftHeader, centerHeader, rightHeader, leftFooter, centerFooterFirst, centerFooterSecond, rightFooter, ranges.printArea);
                    //}

                    // progress update
                    progressSheets += incr;
                    ((IProgress<int>)progress).Report((int)progressSheets);
                }

                // delete default first sheet on creation of workbook
                Wb2Print.Activate();
                Wb2Print.Worksheets[1].Delete();

                // print sheets
                string outputFile = outputPath + @"\catalog.pdf";
                if (File.Exists(outputFile) && ExcelUtility.IsFileInUse(outputFile))
                    throw new Exception(outputFile + " is open, please close it and press 'Print' again.");
                if (!Directory.Exists(outputPath))
                    Directory.CreateDirectory(outputPath);

                // progress update
                ((IProgress<int>)progress).Report(100);

                Wb2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile, OpenAfterPublish: true);

                ExcelUtility.CloseWorkbook(MasterWb, false);
                ExcelUtility.CloseWorkbook(Wb2Print, true);

                var source = new FileInfo(_tmpWorkbookDir + @"\" + _tmpWokbookName);
                source.CopyTo(_tmpWorkbookDir + @"\_" + _tmpWokbookName, true);
                File.Delete(_tmpWorkbookDir + @"\" + _tmpWokbookName);

            }
            catch (Exception ex)
            {
                ExcelUtility.CloseWorkbook(MasterWb, false);
                ExcelUtility.CloseWorkbook(Wb2Print, true);
                //File.Delete(_tmpWorkbookDir + _tmpWokbookName);

                ExcelUtility.CloseExcel();
                throw new Exception(ex.Message);
            }
        }

        public static void FormatSheet(Worksheet sh, string leftHeader, string centerHeader, string rightHeader, string leftFooter, string centerFooterFirst, string centerFooterSecond, string rightFooter, string printArea)
        {
            sh.PageSetup.LeftHeader = "&\"Arial\"&12" + leftHeader;
            sh.PageSetup.CenterHeader = "&\"Arial\"&12" + "&P/&N";
            sh.PageSetup.RightHeader = "&\"Arial\"&12 " + rightHeader;
            sh.PageSetup.LeftFooter = "&\"Arial\"&12 " + leftFooter;
            sh.PageSetup.CenterFooter = "&B&\"Arial\"&16" + centerFooterFirst + "\n" + centerFooterSecond + "&B";
            sh.PageSetup.RightFooter = "&\"Arial\"&12" + rightFooter;

            sh.PageSetup.PrintArea = printArea;

            sh.PageSetup.Zoom = false;
            sh.PageSetup.FitToPagesWide = 1;
            sh.PageSetup.FitToPagesTall = 1;
            sh.PageSetup.CenterVertically = true;
            sh.PageSetup.CenterHorizontally = true;

            sh.PageSetup.LeftMargin = ExcelUtility.XlApp.InchesToPoints(0.5);
            sh.PageSetup.RightMargin = ExcelUtility.XlApp.InchesToPoints(0.5);
            sh.PageSetup.TopMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.BottomMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.HeaderMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
            sh.PageSetup.FooterMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
        }
    }
}
