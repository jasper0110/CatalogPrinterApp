using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtil
{
    public static class ConverterUtility
    {

        public static void NullCheck(string str, string path, string varName)
        {
            if (str == null)
                throw new Exception($"Could not find '{varName}' key in config file {path} !");
        }

        public static List<string> Range2List(string range)
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

        public static Dictionary<int, string> MultipleRange2List(string sheetInput)
        {
            var sheets2Print = new Dictionary<int, string>();

            var index = 0;
            var sheets = sheetInput.Split(';').ToList();
            foreach (var sheet in sheets)
            {
                if (sheet.Contains("-"))
                {
                    foreach (var str in Range2List(sheet))
                    {
                        sheets2Print.Add(index, str);
                        index++;
                    }
                }
                else
                {
                    if (sheet.Length > 0)
                    {
                        sheets2Print.Add(index, sheet);
                        index++;
                    }
                }
            }
            return sheets2Print;
        }

        public static KeyValuePair<int, int> StringRange2Coordinate(string range, string varName)
        {
            var column = new string(range.TakeWhile(char.IsUpper).ToArray());
            if (column.Length < 1)
                return new KeyValuePair<int, int>(0, 0);

            var row = range.Substring(column.Length);

            int columnInt = 0;
            foreach (char c in column)
            {
                columnInt += char.ToUpper(c) - 64;
            }

            int rowInt = 0;
            if (!Int32.TryParse(row, out rowInt))
                throw new Exception($"Parse exception for row " + row + " from range " + range + "!");

            if (rowInt == 0)
                throw new Exception($"Invalid input for cell range '{varName}' : {range} !");

            return new KeyValuePair<int, int>(rowInt, columnInt);
        }

        public static int StringCatalogType2Int(string catalogType)
        {
            int catalogTypeInt = 0;
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
            return catalogTypeInt;
        }
    }
}
