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
        public bool keepExports;
        public List<string> landscapePages;
    }

    public static class ExcelUtility
    {
        public static readonly string _tmpWorkbookDir = @"C:\ProgramData\CatalogPrinter\exports";

        private static Application OpenApplication()
        {
            var excel = new Application();
            excel.DisplayAlerts = false;
            excel.Visible = false;
            return excel;
        }

        public static Application XlApp { get; set; } = null;
        public static Workbook MasterWb { get; set; } = null;
        public static Workbook Wb2Print { get; set; } = null;

        public static void ChangePassword(string wbFullName, string oldPassword, string newPassword)
        {
            Workbook wb = ExcelUtility.GetWorkbook(wbFullName, false, oldPassword);
            if (wb == null)
                throw new Exception($"Wrong password for workbook " + wbFullName + "!");
            wb.Password = newPassword;
            ExcelUtility.CloseWorkbook(wb, true);
        }

        public static Workbook GetWorkbook(string fullName, bool readOnly, string password = null)
        {
            try
            {
                if (password != null)
                    return XlApp?.Workbooks?.Open(fullName, Password: password, ReadOnly: true);

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
                XlApp = null;
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

        public static Dictionary<int, string> GetSheetOrder(AppParameters parameters, string catalogType)
        {
            var sheetOrder = new List<string>();

            if (ExcelUtility.GetWorksheetByName(MasterWb, parameters.sheetSummaryName) == null)
                throw new Exception($"Sheet " + parameters.sheetSummaryName + " not found in workbook " + MasterWb + "!" +
                    "\nPlease check the name of the summary sheet.");

            var summarySheet = MasterWb.Sheets[parameters.sheetSummaryName];

            var header = summarySheet.Rows("1:1").Item[1].Value;
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
            var sheetRank = summarySheet.Columns(rankRange).Item[1].Value;
            var sheetName = summarySheet.Columns("A:A").Item[1].Value;

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

            return sheetOrderDict.ToDictionary(x => x.Key, x => x.Value);
        }

        public static void ExportWorkbook2Pdf(IProgress<int> progress, 
            AppParameters parameters, 
            string password, 
            string catalogType, 
            Dictionary<int, string> sheetOrder,
            bool inclBtw,
            int korting,
            bool printTarieven)
        {
            try
            {
                // open application
                if (XlApp == null)
                    XlApp = OpenApplication();

                // open master workbook
                MasterWb = GetWorkbook(parameters.masterCatalog, true, password);

                // full catalog or selection
                bool printFullCatalog = sheetOrder == null;

                var allSheetOrder = GetSheetOrder(parameters, catalogType);
                // get correct sheet order
                if (!printTarieven)
                {
                    if(!printFullCatalog)
                    {
                        var sheetsToPrint = sheetOrder;
                        if (sheetsToPrint.Select(x => Int32.Parse(x.Value)).Max() > allSheetOrder.Count)
                        {
                            throw new Exception($"Attempting to print page number {sheetsToPrint.Max()}, " +
                                $"but only found {allSheetOrder.Count} given sheets.");
                        }
                        sheetOrder = allSheetOrder.Where(x => sheetOrder.ContainsValue(x.Value))
                            .ToDictionary(x => x.Key, x => x.Value);
                    }
                    else
                    {
                        sheetOrder = GetSheetOrder(parameters, catalogType);
                    }
                }
                else
                {                    
                    var sheetOrderTemp = allSheetOrder.Where(x => sheetOrder.ContainsValue(x.Value))
                            .ToDictionary(x => x.Key, x => x.Value);
                    foreach(var item in sheetOrder)
                    {
                        if(!sheetOrderTemp.ContainsValue(item.Value))
                            throw new Exception($"Could not find sheet name " + item.Value + " in " + parameters.sheetSummaryName +
                                " for catalog type " + catalogType + " or it has been overwritten." +
                            "\nPlease check the sheet order input.");
                    }
                    sheetOrder = sheetOrderTemp;
                }
                               

                // open temp workbook to which the sheets of interest are copied to
                Wb2Print = XlApp?.Workbooks.Add();
                var exportWorkbookName = $@"export_{DateTime.Now.ToString("yyyyMMdd'T'HH'-'mm'-'ss_fff")}.xlsx";
                if (!Directory.Exists(_tmpWorkbookDir))
                    Directory.CreateDirectory(_tmpWorkbookDir);
                Wb2Print?.SaveAs(_tmpWorkbookDir + @"\" + exportWorkbookName);

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
                foreach (var item in sheetOrder)
                {
                    var shName = item.Value;

                    if (ExcelUtility.GetWorksheetByName(MasterWb, shName) == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + MasterWb + "!" +
                            "\nPlease check the sheet order input.");

                    if (printFullCatalog && (MasterWb.Sheets[shName] as Worksheet).IsDrawingTarief())
                        continue;

                    /*
                    * MasterWb manipulations
                    */
                    // unprotect worksheet
                    MasterWb.Sheets[shName].Unprotect();
                    // unhide all columns
                    MasterWb.Sheets[shName].Cells(1, 1).EntireRow.EntireColumn.Hidden = false;
                    // set catalog type
                    MasterWb.Sheets[shName].Cells[cellCatalogType.Key, cellCatalogType.Value] = catalogTypeInt;
                    // set korting
                    if (korting > 0)
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
                    // copy sheet
                    MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                    /*
                    * Wb2Print manipulations
                    */
                    // get format data                    
                    var fd = new FormatData
                    {
                        leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "",
                        centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "",
                        rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "",
                        leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "",
                        centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "",
                        centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "",
                        rightFooter = "TARIEF Nr. " + shName + " / PG " + item.Key,
                        printArea = parameters.ranges.printArea,
                        landscape = parameters.landscapePages.Contains(shName) ? true : false
                    };
                    //format sheet
                    FormatSheet(Wb2Print.Sheets[shName], fd, catalogTypeInt, korting);

                    // copy duplicate sheet
                    if (!printTarieven
                        && catalogTypeInt == (int)CatalogType.PARTICULIER
                        && !((MasterWb.Sheets[shName] as Worksheet).IsDrawingTarief() || (MasterWb.Sheets[shName] as Worksheet).IsCoverTarief()))
                    {
                        /*
                         * MasterWb manipulations
                         */
                        // set btw false
                        MasterWb.Sheets[shName].Cells[cellBtw.Key, cellBtw.Value] = 2;
                        // copy sheet
                        MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);

                        /*
                         * Wb2Print manipulations
                         */
                        // get format data
                        var fd2 = new FormatData
                        {
                            leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "",
                            centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "",
                            rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "",
                            leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "",
                            centerFooterFirst = (MasterWb.Sheets[shName].Cells[cellFooterMidFirst.Key, cellFooterMidFirst.Value] as Range).Value as string ?? "",
                            centerFooterSecond = (MasterWb.Sheets[shName].Cells[cellFooterMidSecond.Key, cellFooterMidSecond.Value] as Range).Value as string ?? "",
                            rightFooter = "TARIEF Nr. " + shName + " / PG " + item.Key,
                            printArea = parameters.ranges.printArea,
                            landscape = parameters.landscapePages.Contains(shName) ? true : false
                        };
                        //format sheet
                        FormatSheet(Wb2Print.Sheets[shName + " (2)"], fd2, catalogTypeInt, korting);
                    }

                    // progress update
                    progressSheets += incr;
                    progress.Report((int)progressSheets);
                }

                if (sheetOrder.Count > 0)
                {
                    // delete default first sheet on creation of workbook
                    Wb2Print.Activate();
                    Wb2Print.Worksheets[1].Delete();

                    // print sheets
                    string outputFile = parameters.outputPath + @"\catalog.pdf";
                    if (File.Exists(outputFile) && ExcelUtility.IsFileInUse(outputFile))
                        throw new Exception(outputFile + " is open, please close it and press 'Print' again.");
                    if (!Directory.Exists(parameters.outputPath))
                        Directory.CreateDirectory(parameters.outputPath);

                    Wb2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile, OpenAfterPublish: true);
                }
                else
                {
                    throw new Exception("No sheets found to print. Please check the input.");
                }

                // progress update
                progress.Report(100);

                ExcelUtility.CloseWorkbook(MasterWb, false);
                ExcelUtility.CloseWorkbook(Wb2Print, true);

                if(!parameters.keepExports)
                    File.Delete(_tmpWorkbookDir + @"\" + exportWorkbookName);

            }
            catch (Exception ex)
            {
                if(MasterWb != null)
                    CloseWorkbook(MasterWb, false);
                if(Wb2Print != null)
                    CloseWorkbook(Wb2Print, true);
   
                CloseExcel();
                throw new Exception(ex.Message);
            }
        }

        public static void FormatSheet(Worksheet sh, FormatData d, int catalogType, int korting)
        {
            if(d.landscape)
                sh.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            else
                sh.PageSetup.Orientation = XlPageOrientation.xlPortrait;

            //if(!sh.IsCoverTarief())
            //    sh.PageSetup.CenterHeader = "&\"Arial\"&12" + "&P";
            sh.PageSetup.CenterHeader = "";

            if (!sh.IsDrawingTarief() && !sh.IsCoverTarief())
            {
                sh.PageSetup.LeftHeader = "&\"Arial\"&12" + d.leftHeader;
                sh.PageSetup.RightHeader = "&\"Arial\"&12 " + d.rightHeader;
                sh.PageSetup.RightFooter = "&\"Arial\"&12" + d.rightFooter;
                sh.PageSetup.LeftFooter = "&\"Arial\"&12 " + d.leftFooter;

                if (catalogType == (int)CatalogType.PARTICULIER && korting <= 0)
                {
                    sh.PageSetup.CenterFooter = "";
                }
                else
                {
                    sh.PageSetup.CenterFooter = "&B&\"Arial\"&16" + d.centerFooterFirst + "\n" + d.centerFooterSecond + "&B";
                }
            }
            else
            {
                sh.PageSetup.LeftHeader = "";
                sh.PageSetup.RightHeader = "";
                sh.PageSetup.LeftFooter = "";
                sh.PageSetup.CenterFooter = "";
                sh.PageSetup.RightFooter = "";
            }

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
