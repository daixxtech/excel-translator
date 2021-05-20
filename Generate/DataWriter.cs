using ExcelTranslator.Extensions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelTranslator.Generate {
    public static class DataWriter {
        /// <summary> 从 DataTable 生成 JSON 文件 </summary>
        public static string DataTableToJSON(DataTable dataTable, Options options) {
            if (!DataUtil.IsValidDataTable(dataTable)) {
                return null;
            }
            var dict = new Dictionary<object, object>();
            var nameDict = DataUtil.GetColumnNameDict(dataTable, NamingStyle.lowerCamel);
            var typeDict = DataUtil.GetColumnTypeDict(dataTable);
            DataColumn idCol = dataTable.Columns[0];
            for (int i = 3, rowCount = dataTable.Rows.Count; i < rowCount; i++) {
                /* 获取对象 ID */
                DataRow row = dataTable.Rows[i];
                string id = row[idCol].ToString();
                if (string.IsNullOrEmpty(id)) {
                    continue;
                }
                /* 构建数据对象 */
                var obj = new Dictionary<string, object>();
                for (int j = 0, colCount = dataTable.Columns.Count; j < colCount; j++) {
                    if (!nameDict.TryGetValue(j, out var name) || !typeDict.TryGetValue(j, out var type)) {
                        continue;
                    }
                    object value = row[dataTable.Columns[j]];
                    obj[name] = type switch {
                        "bool" => (string) value == "true",
                        "int" => (int) (double) value,
                        "float" => value,
                        "string" => value,
                        "bool[]" => ((string) value).Split(",").Select(cell => cell == "true"),
                        "int[]" => ((string) value).Split(",").Select(int.Parse),
                        "float[]" => ((string) value).Split(",").Select(float.Parse),
                        "string[]" => ((string) value).Split(","),
                        var _ => throw new Exception($"Unsupported type: {type}")
                    };
                }
                dict[id] = obj;
            }
            return JsonConvert.SerializeObject(dict);
        }
    }
}
