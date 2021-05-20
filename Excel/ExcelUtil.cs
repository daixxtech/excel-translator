using System;
using System.IO;

namespace ExcelTranslator.Excel {
    public static class ExcelUtil {
        /// <summary> 是否为合法的 Excel 文件（路径、格式） </summary>
        public static bool IsValidExcelFile(string excelPath) {
            if (File.Exists(excelPath)) {
                return IsSupported(excelPath);
            }
            Console.WriteLine("Cannot find excel. Excel path: {0}", excelPath);
            return false;
        }

        /// <summary> 是否为支持的 Excel 文件（格式） </summary>
        public static bool IsSupported(string excelPath) {
            if (excelPath.EndsWith(".xlsx")) {
                return true;
            }
            Console.WriteLine("Only support excel with .xlsx extension. Excel name: {0}", Path.GetFileName(excelPath));
            return false;
        }

        /// <summary> 是否忽略该表格 </summary>
        public static bool IsSheetIgnored(string sheetName) {
            return sheetName.StartsWith("Ignore");
        }

        /// <summary> 是否为枚举表格 </summary>
        public static bool IsEnumSheet(string sheetName) {
            return sheetName.StartsWith("Enum");
        }

        /// <summary> 是否忽略该列 </summary>
        public static bool IsColumnIgnored(string columnName) {
            return columnName.StartsWith("Ignore");
        }
    }
}
