using ExcelTranslator.Extensions;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelTranslator.Generate {
    public struct EnumMember {
        public string name;
        public string value;
        public string comment;
    }

    public struct ClassField {
        public string name;
        public string type;
        public string comment;
    }

    public static class CodeWriter {
        /// <summary> 从 DataTable 生成 C# 代码 </summary>
        public static string DataTableToCSharp(DataTable dataTable, string excelName, Options options) {
            if (!IsValidDataTable(dataTable, options)) {
                return null;
            }
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("/* Auto generated code */").AppendLine();
            builder.AppendFormat("namespace {0} {{", options.ClassNamespace).AppendLine();
            /* 是否需要生成枚举类 */
            List<EnumMember> members = GetEnumMemberList(dataTable);
            if (members != null) {
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public enum E{0} {{", dataTable.TableName).AppendLine();
                foreach (var member in members) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", member.comment).AppendLine();
                    builder.AppendFormat("        {0} = {1},", member.name, member.value).AppendLine();
                }
                builder.AppendLine("    }").AppendLine();
            }
            /* 生成数据类 */
            List<ClassField> fields = GetClassFieldList(dataTable, options);
            builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
            string className = string.IsNullOrEmpty(options.ClassNamePrefix) ? dataTable.TableName : options.ClassNamePrefix + dataTable.TableName;
            builder.AppendFormat("    public class {0} {{", className).AppendLine();
            foreach (var field in fields) {
                builder.AppendFormat("        /// <summary> {0} </summary>", field.comment).AppendLine();
                builder.AppendFormat("        public readonly {0} {1};", field.type, field.name).AppendLine();
            }
            builder.AppendLine("    }");
            builder.AppendLine("}");
            builder.AppendLine().AppendLine("/* End of auto generated code */");
            return builder.ToString();
        }

        /// <summary> 是否为合法的 DataTable </summary>
        private static bool IsValidDataTable(DataTable dataTable, Options options) {
            if (!string.IsNullOrEmpty(options.ExcludePrefix) && dataTable.TableName.StartsWith(options.ExcludePrefix)) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
        }

        /// <summary> 获取对象类的枚举数组 </summary>
        private static List<EnumMember> GetEnumMemberList(DataTable dataTable) {
            if (dataTable.Columns.Count < 3 || dataTable.Rows[0][dataTable.Columns[1]].ToString() != "Enum") {
                return null;
            }
            List<EnumMember> member = new List<EnumMember>();
            DataColumn idCol = dataTable.Columns[0];
            DataColumn nameCol = dataTable.Columns[1];
            DataColumn commentCol = dataTable.Columns[2];
            int rowCount = dataTable.Rows.Count;
            for (int i = 3; i < rowCount; i++) {
                DataRow row = dataTable.Rows[i];
                string name = row[nameCol].ToString();
                if (string.IsNullOrEmpty(name)) {
                    continue;
                }
                member.Add(new EnumMember {
                    name = name.ToNamingStyle(NamingStyle.UpperCamel), value = row[idCol].ToString(), comment = row[commentCol].ToString()
                });
            }
            return member;
        }

        /// <summary> 获取对象类的字段数组 </summary>
        private static List<ClassField> GetClassFieldList(DataTable dataTable, Options options) {
            string excludePrefix = options.ExcludePrefix;
            bool filter = !string.IsNullOrEmpty(excludePrefix);
            List<ClassField> fields = new List<ClassField>();
            DataRow nameRow = dataTable.Rows[0];
            DataRow typeRow = dataTable.Rows[1];
            DataRow commentRow = dataTable.Rows[2];
            foreach (DataColumn col in dataTable.Columns) {
                string name = nameRow[col].ToString();
                if (string.IsNullOrEmpty(name) || filter && name.StartsWith(excludePrefix) || name == "Enum") {
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
