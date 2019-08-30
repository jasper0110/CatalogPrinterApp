using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtil
{
    public static class WorksheetExtensions
    {
        public static bool IsDrawingTarief(this Worksheet ws)
        {
            if (ws == null)
                return false;

            var name = ws.Name;
            int.TryParse(name, out int i);
            return i > 500 && i < 1000;
        }

        public static bool IsCoverTarief(this Worksheet ws)
        {
            if (ws == null)
                return false;

            var name = ws.Name;
            int.TryParse(name, out int i);
            return i >= 1000 && i != 1970 && i != 1980 && i != 1990;
        }
    }
}
