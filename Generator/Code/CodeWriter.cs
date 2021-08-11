using ExcelTranslator.Excel;
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
            ESheet sheetType = ExcelUtil.GetSheetType(dataTable.TableName);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("/* Auto generated code */");
            builder.AppendLine();
            switch (sheetType) {
            case ESheet.EnumSheet: // 生成枚举类
                string enumName = options.EnumNamePrefix + dataTable.TableName.Substring(4);
                List<EnumMember> enumMembers = CodeUtil.GetEnumMembers(dataTable);
                builder.AppendFormat("namespace {0} {{", options.Namespace).AppendLine();
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public enum {0} {{", enumName).AppendLine();
                // 枚举成员
                foreach (var member in enumMembers) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", member.comment).AppendLine();
                    builder.AppendFormat("        {0} = {1},", member.name, member.value).AppendLine();
                }
                builder.AppendLine("    }");
                builder.AppendLine("}");
                break;
            case ESheet.ParamSheet: // 生成参数类 
                string paramName = options.ParamNamePrefix + dataTable.TableName.Substring(5);
                List<ParamField> paramFields = CodeUtil.GetParamFields(dataTable);
                builder.AppendFormat("namespace {0} {{", options.Namespace).AppendLine();
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public static class {0} {{", paramName).AppendLine();
                // 静态只读字段
                foreach (var field in paramFields) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", field.comment).AppendLine();
                    string value = CodeUtil.DataTableValueToCodeForm(field.type, field.value);
                    builder.AppendFormat("        public static readonly {0} {1} = {2};", field.type, field.name, value).AppendLine();
                }
                builder.AppendLine("    }");
                builder.AppendLine("}");
                break;
            case ESheet.ClassSheet: // 生成数据类
                string className = options.ClassNamePrefix + dataTable.TableName;
                List<ClassField> classFields = CodeUtil.GetClassFields(dataTable);
                builder.AppendFormat("namespace {0} {{", options.Namespace).AppendLine();
                builder.AppendFormat("    /// <summary> Generate From {0} </summary>", excelName).AppendLine();
                builder.AppendFormat("    public class {0} {{", className).AppendLine();
                // 字段
                foreach (var field in classFields) {
                    builder.AppendFormat("        /// <summary> {0} </summary>", field.comment).AppendLine();
                    builder.AppendFormat("        public readonly {0} {1};", field.type, field.name).AppendLine();
                }
                // 构造函数
                builder.AppendLine();
                builder.AppendFormat("        public {0}(", className)
                       .AppendJoin(", ", classFields.Select(field => string.Format("{0} {1}", field.type, field.name)))
                       .AppendLine(") {");
                foreach (var field in classFields) {
                    builder.AppendFormat("            this.{0} = {0};", field.name).AppendLine();
                }
                builder.AppendLine("        }");
                builder.AppendLine("    }");
                builder.AppendLine("}");
                break;
            }
            builder.AppendLine();
            builder.AppendLine("/* End of auto generated code */");
            return builder.ToString();
        }
    }
}
