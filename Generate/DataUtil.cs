using ExcelTranslator.Extensions;
using System.Collections.Generic;
using System.Data;

namespace ExcelTranslator.Generate {
    public static class DataUtil {
        /// <summary> 是否为合法的 DataTable </summary>
        public static bool IsValidDataTable(DataTable dataTable) {
            if (dataTable.TableName.StartsWith("Exclude") || dataTable.TableName.EndsWith("Def")) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
        }

        /// <summary> 获取 DataTable 的每列名称字典，键：下标，值：名称 </summary>
        public static Dictionary<int, string> GetColumnNameDict(DataTable dataTable, NamingStyle style) {
            var nameDict = new Dictionary<int, string>();
            int colCount = dataTable.Columns.Count;
            DataRow nameRow = dataTable.Rows[0];
            for (int i = 0; i < colCount; i++) {
                string name = nameRow[dataTable.Columns[i]].ToString();
                if (string.IsNullOrEmpty(name) || name.StartsWith("Exclude")) {
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
