using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelTranslator.Generator.Code {
    public static class CodeWriter {
        /// <summary> 从 DataTable 生成 C# 代码 </summary>
        public static string DataTableToCSharp(DataTable dataTable, string excelName, Options options) {
            if (!CodeUtil.IsValidDataTable(dataTable)) {
                return null;
            }
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("/* Auto generated code */");
            builder.AppendLine();
            builder.AppendFormat("namespace {0} {{", options.ClassNamespace).AppendLine();

            List<EnumMember> members = CodeUtil.GetEnumMembers(dataTable);
            if (members != null) {
                /* 生成枚举类 */
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public enum E{0} {{", dataTable.TableName).AppendLine();
                // 枚举成员
                foreach (var member in members) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", member.comment).AppendLine();
                    builder.AppendFormat("        {0} = {1},", member.name, member.value).AppendLine();
                }
                builder.AppendLine("    }");
            } else {
                /* 生成数据类 */
                string className = options.ClassNamePrefix + dataTable.TableName;
                List<ClassField> fields = CodeUtil.GetClassFields(dataTable, options);
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public class {0} {{", className).AppendLine();
                // 字段
                foreach (var field in fields) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", field.comment).AppendLine();
                    builder.AppendFormat("        public readonly {0} {1};", field.type, field.name).AppendLine();
                }
                builder.AppendLine();
                // 构造函数
                builder.AppendFormat("        public {0}(", className).AppendJoin(", ", fields.Select(field => $"{field.type} {field.name}")).AppendLine("){");
                foreach (var field in fields) {
                    builder.AppendFormat("            this.{0} = {0};", field.name).AppendLine();
                }
                builder.AppendLine("        }");
                builder.AppendLine("    }");
            }

            builder.AppendLine("}");
            builder.AppendLine();
            builder.AppendLine("/* End of auto generated code */");
            return builder.ToString();
        }
    }
}
