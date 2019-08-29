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
        public string korting;
        public string footerRight;
        public string footerLeft;
        public string footerMidFirst;
        public string footerMidSecond;
        public string headerRight;
        public string headerLeft;
        public string headerMid;
        public string printArea;
    }

    public struct AppParameters
    {
        public InputRanges ranges;
        public string masterCatalog;
        public string outputPath;
        public string hash;
        public string sheetSummaryName;
    }

    public static class ExcelUtility
    {
        public static readonly string _tmpWorkbookDir = @"C:\temp";
        public static readonly string _tmpWokbookName = @"temp.xlsx";

        private static Application OpenApplication()
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

        public static Application XlApp { get; set; } = null;
        public static Workbook MasterWb { get; set; } = null;
        public static Workbook Wb2Print { get; set; } = null;

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
                    return XlApp?.Workbooks?.Open(fullName, Password: password);

                return XlApp?.Workbooks?.Open(fullName);
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
            if (XlApp != null)
            {
                XlApp.Quit();
                Marshal.ReleaseComObject(XlApp);
            }
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

        public static List<string> GetSheetOrder(AppParameters parameters, string catalogType)
        {
            var sheetOrder = new List<string>();

            if (ExcelUtility.GetWorksheetByName(MasterWb, parameters.sheetSummaryName) == null)
                throw new Exception($"Sheet " + parameters.sheetSummaryName + " not found in workbook " + MasterWb + "!" +
                    "\nPlease check the name of the summary sheet.");
            var header = MasterWb.Sheets[parameters.sheetSummaryName].Rows("1:1").Item[1].Value;
            int columnInt = -1;
            for (int i = 1; i <= 100; ++i)
            {
                var str = header[1, i];
                if (str == catalogType)
                {
                    columnInt = i;
                    break;
                }
            }

            if (columnInt < 0)
                throw new Exception($"Catalog type " + catalogType + " not found in worksheet " + parameters.sheetSummaryName + "!" +
                    "\nPlease check the name of the summary sheet and if the catalog type exists in the sheet.");

            char columnChar = (char)(columnInt + 64);
            string rankRange = columnChar + ":" + columnChar;
            var sheetRank = MasterWb.Sheets[parameters.sheetSummaryName].Columns(rankRange).Item[1].Value;
            var sheetName = MasterWb.Sheets[parameters.sheetSummaryName].Columns("A:A").Item[1].Value;

            var sheetOrderDict = new SortedDictionary<int, string>();
            for (int i = 2; i < 1000; ++i)
            {
                var rank = Convert.ToInt32(sheetRank[i, 1]);
                var name = Convert.ToString(sheetName[i, 1]);
                // check if item already exists
                // add parse safety for rank and name
                if (rank > 0 && name != null)
                    sheetOrderDict[rank] = name;
            }
            sheetOrder = sheetOrderDict.Values.ToList();

            return sheetOrder;
        }

        public static void ExportWorkbook2Pdf(IProgress<int> progress, 
            AppParameters parameters, 
            string password, 
            string catalogType, 
            List<string> sheetOrder,
            bool inclBtw,
            int korting,
            bool printTarieven = true)
        {
            try
            {
                // open application
                if (XlApp == null)
                    XlApp = OpenApplication();

                // open master workbook
                MasterWb = GetWorkbook(parameters.masterCatalog, password);

                // get correct sheet order
                if(!printTarieven)
                {
                    if(sheetOrder != null)
                    {
                        var sheetsToPrint = sheetOrder.Select(x => Int32.Parse(x)).ToList();                        
                        var allSheetOrder = GetSheetOrder(parameters, catalogType);
                        if (sheetsToPrint.Max() > allSheetOrder.Count)
                        {
                            throw new Exception($"Attempting to print page number {sheetsToPrint.Max()}, " +
                                $"but only found {allSheetOrder.Count} given sheets.");
                        }
                        sheetOrder = sheetsToPrint.Select(x => allSheetOrder[x-1]).ToList();
                    }
                    else
                    {
                        sheetOrder = GetSheetOrder(parameters, catalogType);
                    }
                }
                               

                // open temp workbook to which the sheets of interest are copied to
                Wb2Print = XlApp?.Workbooks.Add();
                if (!Directory.Exists(_tmpWorkbookDir))
                    Directory.CreateDirectory(_tmpWorkbookDir);
                Wb2Print?.SaveAs(_tmpWorkbookDir + @"\" + _tmpWokbookName);

                // progress update
                progress.Report(30);

                var cellFooterRight = ConverterUtility.StringRange2Coordinate(parameters.ranges.footerRight, nameof(parameters.ranges.footerRight));
                var cellFooterLeft = ConverterUtility.StringRange2Coordinate(parameters.ranges.footerLeft, nameof(parameters.ranges.footerLeft));
                var cellFooterMidFirst = ConverterUtility.StringRange2Coordinate(parameters.ranges.footerMidFirst, nameof(parameters.ranges.footerMidFirst));
                var cellFooterMidSecond = ConverterUtility.StringRange2Coordinate(parameters.ranges.footerMidSecond, nameof(parameters.ranges.footerMidSecond));
                var cellHeaderRight = ConverterUtility.StringRange2Coordinate(parameters.ranges.headerRight, nameof(parameters.ranges.headerRight));
                var cellHeaderLeft = ConverterUtility.StringRange2Coordinate(parameters.ranges.headerLeft, nameof(parameters.ranges.headerLeft));
                var cellHeaderMid = ConverterUtility.StringRange2Coordinate(parameters.ranges.headerMid, nameof(parameters.ranges.headerMid));

                var cellCatalogType = ConverterUtility.StringRange2Coordinate(parameters.ranges.catalogType, nameof(parameters.ranges.catalogType));
                var cellKorting = ConverterUtility.StringRange2Coordinate(parameters.ranges.korting, nameof(parameters.ranges.korting));
                var cellBtw = ConverterUtility.StringRange2Coordinate(parameters.ranges.btw, nameof(parameters.ranges.btw));

                // catalog int type
                int catalogTypeInt = ConverterUtility.StringCatalogType2Int(catalogType);                

                double progressSheets = 30.0;
                double incr = 70.0 / (double)sheetOrder.Count;

                // copy necessary sheets to temp workbook and put sheets in correct order
                foreach (var shName in sheetOrder)
                {
                    if (ExcelUtility.GetWorksheetByName(MasterWb, shName) == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + MasterWb + "!" +
                            "\nPlease check the sheet order input.");

                    // unprotect worksheet
                    MasterWb.Sheets[shName].Unprotect();

                    // set catalog type
                    MasterWb.Sheets[shName].Cells[cellCatalogType.Key, cellCatalogType.Value] = catalogTypeInt;
                    // set korting
                    MasterWb.Sheets[shName].Cells[cellKorting.Key, cellKorting.Value] = korting;
                    // set btw
                    if (inclBtw || (!printTarieven && catalogTypeInt == (int)CatalogType.PARTICULIER))
                    {
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 1;
                    }
                    else
                    {
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 2;
                    }

                    // get format data
                    string leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "";
                    string centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "";
                    string rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "";
                    string leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "";
                    string centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "";
                    string centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "";
                    string rightFooter = "TARIEF Nr. " + shName;
                    var fd = new FormatData
                    {
                        leftHeader = leftHeader,
                        centerHeader = centerHeader,
                        rightHeader = rightHeader,
                        leftFooter = leftFooter,
                        centerFooterFirst = centerFooterFirst,
                        centerFooterSecond = centerFooterSecond,
                        rightFooter = rightFooter,
                        printArea = parameters.ranges.printArea
                    };

                    // unhide all columns
                    MasterWb.Sheets[shName].Cells(1, 1).EntireRow.EntireColumn.Hidden = false;

                    // copy sheet
                    MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                    //format sheet
                    FormatSheet(Wb2Print.Sheets[shName], fd);

                    // copy sheet
                    if (!printTarieven && catalogTypeInt == (int)CatalogType.PARTICULIER)
                    {
                        // set btw false
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 2;

                        // get format data
                        leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "";
                        centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "";
                        rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "";
                        leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "";
                        centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "";
                        centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "";
                        rightFooter = "TARIEF Nr. " + shName;
                        var fd2 = new FormatData
                        {
                            leftHeader = leftHeader,
                            centerHeader = centerHeader,
                            rightHeader = rightHeader,
                            leftFooter = leftFooter,
                            centerFooterFirst = centerFooterFirst,
                            centerFooterSecond = centerFooterSecond,
                            rightFooter = rightFooter,
                            printArea = parameters.ranges.printArea
                        };

                        // copy sheet
                        MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                        //format sheet
                        FormatSheet(Wb2Print.Sheets[shName + " (2)"], fd2);
                    }

                    // progress update
                    progressSheets += incr;
                    progress.Report((int)progressSheets);
                }

                // delete default first sheet on creation of workbook
                Wb2Print.Activate();
                Wb2Print.Worksheets[1].Delete();

                // print sheets
                string outputFile = parameters.outputPath + @"\catalog.pdf";
                if (File.Exists(outputFile) && ExcelUtility.IsFileInUse(outputFile))
                    throw new Exception(outputFile + " is open, please close it and press 'Print' again.");
                if (!Directory.Exists(parameters.outputPath))
                    Directory.CreateDirectory(parameters.outputPath);

                // progress update
                progress.Report(100);

                Wb2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile, OpenAfterPublish: true);

                ExcelUtility.CloseWorkbook(MasterWb, false);
                ExcelUtility.CloseWorkbook(Wb2Print, true);

                var source = new FileInfo(_tmpWorkbookDir + @"\" + _tmpWokbookName);
                source.CopyTo(_tmpWorkbookDir + @"\_" + _tmpWokbookName, true);
                File.Delete(_tmpWorkbookDir + @"\" + _tmpWokbookName);

            }
            catch (Exception ex)
            {
                if(MasterWb != null)
                    CloseWorkbook(MasterWb, false);
                if(Wb2Print != null)
                    CloseWorkbook(Wb2Print, true);
                //File.Delete(_tmpWorkbookDir + _tmpWokbookName);

                CloseExcel();
                throw new Exception(ex.Message);
            }
        }

        public static void FormatSheet(Worksheet sh, FormatData d)
        {
            sh.PageSetup.LeftHeader = "&\"Arial\"&12" + d.leftHeader;
            sh.PageSetup.CenterHeader = "&\"Arial\"&12" + "&P/&N";
            sh.PageSetup.RightHeader = "&\"Arial\"&12 " + d.rightHeader;
            sh.PageSetup.LeftFooter = "&\"Arial\"&12 " + d.leftFooter;
            sh.PageSetup.CenterFooter = "&B&\"Arial\"&16" + d.centerFooterFirst + "\n" + d.centerFooterSecond + "&B";
            sh.PageSetup.RightFooter = "&\"Arial\"&12" + d.rightFooter;

            sh.PageSetup.PrintArea = d.printArea;

            sh.PageSetup.Zoom = false;
            sh.PageSetup.FitToPagesWide = 1;
            sh.PageSetup.FitToPagesTall = 1;
            sh.PageSetup.CenterVertically = true;
            sh.PageSetup.CenterHorizontally = true;

            sh.PageSetup.LeftMargin = XlApp?.InchesToPoints(0.5) ?? 0.0;
            sh.PageSetup.RightMargin = XlApp?.InchesToPoints(0.5) ?? 0.0;
            sh.PageSetup.TopMargin = XlApp?.InchesToPoints(0.7) ?? 0.0;
            sh.PageSetup.BottomMargin = XlApp?.InchesToPoints(0.7) ?? 0.0;
            sh.PageSetup.HeaderMargin = XlApp?.InchesToPoints(0.3) ?? 0.0;
            sh.PageSetup.FooterMargin = XlApp?.InchesToPoints(0.3) ?? 0.0;
        }
    }
}
