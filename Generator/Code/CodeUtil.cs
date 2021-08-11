using ExcelTranslator.Excel;
using ExcelTranslator.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelTranslator.Generator.Code {
    public static class CodeUtil {
        /// <summary> 是否为合法的 DataTable </summary>
        public static bool IsValidDataTable(DataTable dataTable) {
            if (ExcelUtil.IsSheetIgnored(dataTable.TableName)) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
        }

        /// <summary> DataTable 的值转换为代码形式 </summary>
        public static string DataTableValueToCodeForm(string type, object value) {
            return type switch {
                "bool" => value is DBNull ? "false" : (string) value,
                "int" => value is DBNull ? "0" : ((int) (double) value).ToString(),
                "float" => value is DBNull ? "0.0F" : string.Format("{0}F", value),
                "string" => value is DBNull ? "\"\"" : string.Format("\"{0}\"", value),
                "bool[]" => value is DBNull
                    ? null
                    : string.Format("{{{0}}}", string.Join(", ", ((string) value).Split(",").Select(cell => cell.Trim() == "true" ? "true" : "false"))),
                "int[]" => value is DBNull
                    ? null
                    : string.Format("{{{0}}}", string.Join(", ", ((string) value).Split(",").Select(str => int.Parse(str.Trim()).ToString()))),
                "float[]" => value is DBNull
                    ? null
                    : string.Format("{{{0}}}", string.Join(", ", ((string) value).Split(",").Select(str => string.Format("{0}F", float.Parse(str.Trim()))))),
                "string[]" => value is DBNull
                    ? null
                    : string.Format("{{{0}}}", string.Join(", ", ((string) value).Split(",").Select(str => string.Format("\"{0}\"", str)))),
                _ => throw new Exception($"Unsupported type: {type}")
            };
        }

        /// <summary> 获取枚举表的成员数组 </summary>
        public static List<EnumMember> GetEnumMembers(DataTable dataTable) {
            var members = new List<EnumMember>();
            DataColumn enumCol = dataTable.Columns[0];
            DataColumn nameCol = dataTable.Columns[1];
            DataColumn commentCol = dataTable.Columns[2];
            for (int i = 3, rowCount = dataTable.Rows.Count; i < rowCount; i++) {
                DataRow row = dataTable.Rows[i];
                string name = row[nameCol].ToString();
                if (string.IsNullOrEmpty(name)) {
                    continue;
                }
                members.Add(new EnumMember {
                    name = name.ToNamingStyle(NamingStyle.UpperCamel), value = row[enumCol].ToString(), comment = row[commentCol].ToString()
                });
            }
            return members;
        }

        /// <summary> 获取数据表的字段数组 </summary>
        public static List<ClassField> GetClassFields(DataTable dataTable) {
            var fields = new List<ClassField>();
            DataRow nameRow = dataTable.Rows[0];
            DataRow typeRow = dataTable.Rows[1];
            DataRow commentRow = dataTable.Rows[2];
            foreach (DataColumn col in dataTable.Columns) {
                string name = nameRow[col].ToString();
                if (string.IsNullOrEmpty(name) || ExcelUtil.IsColumnIgnored(name)) {
                    continue;
                }
                fields.Add(new ClassField {
                    name = name.ToNamingStyle(NamingStyle.lowerCamel), type = typeRow[col].ToString(), comment = commentRow[col].ToString()
                });
            }
            return fields;
        }

        /// <summary> 获取参数表的字段数组 </summary>
        public static List<ParamField> GetParamFields(DataTable dataTable) {
            var fields = new List<ParamField>();
            DataColumn nameCol = dataTable.Columns[0];
            DataColumn typeCol = dataTable.Columns[1];
            DataColumn valueCol = dataTable.Columns[2];
            DataColumn commentCol = dataTable.Columns[3];
            foreach (DataRow row in dataTable.Rows) {
                string name = row[nameCol].ToString();
                if (string.IsNullOrEmpty(name) || ExcelUtil.IsColumnIgnored(name)) {
                    continue;
                }
                fields.Add(new ParamField {
                    name = name.ToNamingStyle(NamingStyle.lowerCamel),
                    type = row[typeCol].ToString(),
                    value = row[valueCol],
                    comment = row[commentCol].ToString()
                });
            }
            return fields;
        }
    }
}
