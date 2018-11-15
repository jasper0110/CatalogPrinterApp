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
    public struct InputRanges
    {       
        public string catalogType;
        public string btw;
        public string footerRight;
        public string footerLeft;
        public string footerMid;
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
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread 
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

        public static void ExportWorkbook2Pdf(string wbName, string password, string catalogType, string outputPath, List<string> sheetOrder, InputRanges ranges)
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

                string leftHeader = "null", centerHeader = "null", rightHeader = "null", leftFooter = "null", rightFooter = "null";

                var cellFooterRight = StringRange2Coordinate(ranges.footerRight);
                if (cellFooterRight.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterRight: " + ranges.footerRight + "!");
                var cellFooterLeft = StringRange2Coordinate(ranges.footerLeft);
                if (cellFooterLeft.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterLeft: " + ranges.footerLeft + "!");
                var cellFooterMid = StringRange2Coordinate(ranges.footerMid);
                if (cellFooterMid.Key == 0)
                    throw new Exception($"Invalid input for cell range cellFooterMid: " + ranges.footerMid + "!");
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

                // copy necessary sheets to temp workbook and put sheets in correct order
                foreach (var shName in sheetOrder)
                {
                    if (ExcelUtility.GetWorksheetByName(MasterWb, shName) == null)
                        throw new Exception($"Sheet " + shName + " not found in workbook " + MasterWb + "!" +
                            "\nPlease check the sheet order input.");
                    // set catalog type
                    MasterWb.Sheets[shName].Cells[cellCatalogType.Key, cellCatalogType.Value] = catalogType;

                    leftHeader = (MasterWb.Sheets[shName].Cells[cellHeaderLeft.Key, cellHeaderLeft.Value] as Range).Value as string ?? "null";
                    centerHeader = (MasterWb.Sheets[shName].Cells[cellHeaderMid.Key, cellHeaderMid.Value] as Range).Value as string ?? "null";
                    rightHeader = (MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value as string ?? "null";
                    //var rightHeaderDate = ((MasterWb.Sheets[shName].Cells[cellHeaderRight.Key, cellHeaderRight.Value] as Range).Value);
                    //rightHeader = "null";
                    //if (rightHeaderDate != null)
                    //    rightHeader = rightHeaderDate.ToString("dd/MM/yyyy");
                    leftFooter = (MasterWb.Sheets[shName].Cells[cellFooterLeft.Key, cellFooterLeft.Value] as Range).Value as string ?? "null";
                    rightFooter = (MasterWb.Sheets[shName].Cells[cellFooterRight.Key, cellFooterRight.Value] as Range).Value as string ?? "null";

                    // copy sheet
                    if (catalogType.ToUpper() == "PARTICULIER")
                    {
                        //SetBtwField(Workbook.Sheets[shName], true);
                        MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);
                        Wb2Print.Sheets[Wb2Print.Sheets.Count].Cells[cellBtw.Key, cellBtw.Value] = "ja";
                        //SetBtwField(Workbook.Sheets[shName], false);
                        MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);
                        Wb2Print.Sheets[Wb2Print.Sheets.Count].Cells[cellBtw.Key, cellBtw.Value] = "neen";
                    }
                    else
                    {
                        MasterWb.Sheets[shName].Copy(After: Wb2Print.Sheets[Wb2Print.Sheets.Count]);
                    }
                }

                // delete default first sheet on creation of workbook
                Wb2Print.Activate();
                Wb2Print.Worksheets[1].Delete();

                // format and print sheets
                string outputFile = outputPath + @"\catalog.pdf";
                if (ExcelUtility.IsFileInUse(outputFile))
                    throw new Exception(outputFile + " is open, please close it and press 'Print' again.");
                foreach (Worksheet sh in Wb2Print.Worksheets)
                    FormatSheet(sh, leftHeader, centerHeader, rightHeader, leftFooter, rightFooter, ranges.printArea);
                Wb2Print.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFile, OpenAfterPublish: true);
                
            }
            catch (Exception ex)
            {
                ExcelUtility.CloseWorkbook(MasterWb, false);
                ExcelUtility.CloseWorkbook(Wb2Print, true);
                File.Delete(_tmpWorkbookDir + _tmpWokbookName);

                ExcelUtility.CloseExcel();
                throw new Exception(ex.Message);
            }
        }

        public static void FormatSheet(Worksheet sh, string leftHeader, string centerHeader, string rightHeader, string leftFooter, string rightFooter, string printArea)
        {
            sh.PageSetup.LeftHeader = leftHeader;
            sh.PageSetup.CenterHeader = centerHeader;
            sh.PageSetup.RightHeader = rightHeader;
            sh.PageSetup.LeftFooter = leftFooter;
            sh.PageSetup.RightFooter = rightFooter;

            sh.PageSetup.PrintArea = printArea;

            sh.PageSetup.Zoom = false;
            sh.PageSetup.FitToPagesWide = 1;
            sh.PageSetup.FitToPagesTall = 1;
            sh.PageSetup.CenterVertically = true;
            sh.PageSetup.CenterHorizontally = true;

            sh.PageSetup.LeftMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.RightMargin = ExcelUtility.XlApp.InchesToPoints(0.7);
            sh.PageSetup.TopMargin = ExcelUtility.XlApp.InchesToPoints(0.75);
            sh.PageSetup.BottomMargin = ExcelUtility.XlApp.InchesToPoints(0.75);
            sh.PageSetup.HeaderMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
            sh.PageSetup.FooterMargin = ExcelUtility.XlApp.InchesToPoints(0.3);
        }
    }
}
