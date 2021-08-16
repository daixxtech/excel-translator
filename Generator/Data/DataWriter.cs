using ExcelTranslator.Extensions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelTranslator.Generator.Data {
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
                    try {
                        obj[name] = DataUtil.DataTableValueToDataForm(type, value);
                    } catch (Exception exception) {
                        Console.WriteLine("[Error] Exception at sheet {0}, row: {1}, col: {2}", dataTable.TableName, (i + 1).ToString(), (j + 1).ToString());
                        Console.WriteLine(exception);
                        throw;
                    }
                }
                dict[id] = obj;
            }
            return JsonConvert.SerializeObject(dict);
        }
    }
}
