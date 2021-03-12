using System.Collections.Generic;
using System.Data;

namespace ExcelTranslator.Extensions {
    public static class DataTableExtension {
        /// <summary> 获取 DataTable 的列名字典 </summary>
        public static Dictionary<int, string> GetColumnNameDict(this DataTable dataTable, string excludePrefix, NamingStyle style) {
            Dictionary<int, string> nameDict = new Dictionary<int, string>();
            int colCount = dataTable.Columns.Count;
            DataRow nameRow = dataTable.Rows[0];
            bool filter = !string.IsNullOrEmpty(excludePrefix);
            for (int i = 0; i < colCount; i++) {
                string name = nameRow[dataTable.Columns[i]].ToString();
                if (string.IsNullOrEmpty(name) || filter && name.StartsWith(excludePrefix)) {
                    continue;
                }
                nameDict.Add(i, name.ToNamingStyle(style));
            }
            return nameDict;
        }
    }
}
