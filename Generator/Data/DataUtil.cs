using ExcelTranslator.Excel;
using ExcelTranslator.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelTranslator.Generator.Data {
    public static class DataUtil {
        /// <summary> 是否为合法的 DataTable </summary>
        public static bool IsValidDataTable(DataTable dataTable) {
            if (ExcelUtil.IsSheetIgnored(dataTable.TableName) || ExcelUtil.GetSheetType(dataTable.TableName) != ESheet.ClassSheet) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
        }

        /// <summary> DataTable 的值转换为数据形式 </summary>
        public static object DataTableValueToDataForm(string type, object value) {
            return type switch {
                "bool" => value is DBNull ? false : value.ToString()?.ToLower() == "true",
                "int" => value is DBNull ? 0 : (int) (double) value,
                "float" => value is DBNull ? 0.0F : value,
                "string" => value is DBNull ? null : value,
                "bool[]" => value is DBNull ? null : ((string) value).Split(",").Select(cell => cell.Trim() == "true"),
                "int[]" => value is DBNull ? null : ((string) value).Split(",").Select(str => int.Parse(str.Trim())),
                "float[]" => value is DBNull ? null : ((string) value).Split(",").Select(str => float.Parse(str.Trim())),
                "string[]" => value is DBNull ? null : ((string) value).Split(","),
                _ => throw new Exception($"Unsupported type: {type}")
            };
        }

        /// <summary> 获取 DataTable 的每列名称字典，键：下标，值：名称 </summary>
        public static Dictionary<int, string> GetColumnNameDict(DataTable dataTable, NamingStyle style) {
            var nameDict = new Dictionary<int, string>();
            int colCount = dataTable.Columns.Count;
            DataRow nameRow = dataTable.Rows[0];
            for (int i = 0; i < colCount; i++) {
                string name = nameRow[dataTable.Columns[i]].ToString();
                if (string.IsNullOrEmpty(name) || ExcelUtil.IsColumnIgnored(name)) {
                    continue;
                }
                nameDict.Add(i, name.ToNamingStyle(style));
            }
            return nameDict;
        }

        /// <summary> 获取 DataTable 的每列类型字典，键：下标，值：类型 </summary>
        public static Dictionary<int, string> GetColumnTypeDict(DataTable dataTable) {
            var typeDict = new Dictionary<int, string>();
            DataRow nameRow = dataTable.Rows[1];
            for (int i = 0, colCount = dataTable.Columns.Count; i < colCount; i++) {
                string name = nameRow[dataTable.Columns[i]].ToString();
                if (string.IsNullOrEmpty(name)) {
                    continue;
                }
                typeDict.Add(i, name);
            }
            return typeDict;
        }
    }
}
