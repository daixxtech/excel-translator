using ExcelTranslator.Extensions;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelTranslator.Generate {
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
            List<ClassField> fields = GetClassFieldList(dataTable, options);
            builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
            string className = string.IsNullOrEmpty(options.ClassNamePrefix) ? dataTable.TableName : options.ClassNamePrefix + dataTable.TableName;
            builder.AppendFormat("    public class {0} {{", className).AppendLine();
            foreach (var field in fields) {
                builder.AppendFormat("        /// <summary> {0} </summary>", field.comment).AppendLine();
                builder.AppendFormat("        public readonly {0} {1};", field.type, field.name).AppendLine();
            }
            builder.AppendLine("    }");
            builder.AppendLine("}").AppendLine();
            builder.AppendLine("/* End of auto generated code */").AppendLine();
            return builder.ToString();
        }

        /// <summary> 是否为合法的 DataTable </summary>
        private static bool IsValidDataTable(DataTable dataTable, Options options) {
            if (!string.IsNullOrEmpty(options.ExcludePrefix) && dataTable.TableName.StartsWith(options.ExcludePrefix)) {
                return false;
            }
            return dataTable.Columns.Count >= 1 && dataTable.Rows.Count >= 3;
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
                if (string.IsNullOrEmpty(name) || filter && name.StartsWith(excludePrefix)) {
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
