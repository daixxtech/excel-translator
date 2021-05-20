using ExcelTranslator.Extensions;
using ExcelTranslator.Generate.Code;
using System.Collections.Generic;
using System.Data;

namespace ExcelTranslator.Generate {
    public static class CodeUtil {
        /// <summary> 是否为合法的 DataTable </summary>
        public static bool IsValidDataTable(DataTable dataTable) {
            if (dataTable.TableName.StartsWith("Exclude")) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
        }

        /// <summary> 获取对象类的枚举数组 </summary>
        public static List<EnumMember> GetEnumMembers(DataTable dataTable) {
            if (dataTable.Columns.Count < 3 || dataTable.Rows[0][dataTable.Columns[1]].ToString() != "Enum") {
                return null;
            }
            var members = new List<EnumMember>();
            DataColumn idCol = dataTable.Columns[0];
            DataColumn nameCol = dataTable.Columns[1];
            DataColumn commentCol = dataTable.Columns[2];
            for (int i = 3, rowCount = dataTable.Rows.Count; i < rowCount; i++) {
                DataRow row = dataTable.Rows[i];
                string name = row[nameCol].ToString();
                if (string.IsNullOrEmpty(name)) {
                    continue;
                }
                members.Add(new EnumMember {
                    name = name.ToNamingStyle(NamingStyle.UpperCamel), value = row[idCol].ToString(), comment = row[commentCol].ToString()
                });
            }
            return members;
        }

        /// <summary> 获取对象类的字段数组 </summary>
        public static List<ClassField> GetClassFields(DataTable dataTable, Options options) {
            var fields = new List<ClassField>();
            DataRow nameRow = dataTable.Rows[0];
            DataRow typeRow = dataTable.Rows[1];
            DataRow commentRow = dataTable.Rows[2];
            foreach (DataColumn col in dataTable.Columns) {
                string name = nameRow[col].ToString();
                if (string.IsNullOrEmpty(name) || name.StartsWith("Exclude")) {
                    continue;
                }
                fields.Add(new ClassField {
                    name = name.ToNamingStyle(NamingStyle.lowerCamel), type = typeRow[col].ToString(), comment = commentRow[col].ToString()
                });
            }
            return fields;
        }
    }
}
