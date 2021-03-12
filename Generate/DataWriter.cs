using ExcelTranslator.Extensions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelTranslator.Generate {
    public static class DataWriter {
        /// <summary> 从 DataTable 生成 JSON 文件 </summary>
        public static string DataTableToJSON(DataTable dataTable, Options options) {
            JsonSerializerSettings jsonSettings = new JsonSerializerSettings {Formatting = Formatting.Indented};
            Dictionary<object, object> objDict = new Dictionary<object, object>();
            Dictionary<int, string> colNameDict = dataTable.GetColumnNameDict(options.ExcludePrefix, NamingStyle.lowerCamel);
            DataColumn idCol = dataTable.Columns[0];
            int rowCount = dataTable.Rows.Count;
            for (int i = 3; i < rowCount; i++) {
                /* 获取对象 ID */
                DataRow row = dataTable.Rows[i];
                string id = row[idCol].ToString();
                if (string.IsNullOrEmpty(id)) {
                    continue;
                }
                Dictionary<string, object> objValueDict = new Dictionary<string, object>();
                int colCount = dataTable.Columns.Count;
                /* 构建数据对象 */
                for (int j = 0; j < colCount; j++) {
                    if (!colNameDict.TryGetValue(j, out var colName)) {
                        continue;
                    }
                    object value = row[dataTable.Columns[j]];
                    /* 特殊类型处理 */
                    switch (value) {
                    case DBNull _:
                        value = string.Empty;
                        break;
                    case double numericValue when Math.Abs((int) numericValue - numericValue) < 1e-6:
                        value = (int) numericValue;
                        break;
                    }
                    objValueDict[colName] = value;
                }
                objDict[id] = objValueDict;
            }
            return JsonConvert.SerializeObject(objDict, jsonSettings);
        }
    }
}
