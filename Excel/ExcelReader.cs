using ExcelDataReader;
using System.Data;
using System.IO;

namespace ExcelTranslator.Excel {
    public static class ExcelReader {
        /// <summary> 读取 Excel 转换为 DataTable </summary>
        public static DataTableCollection ReadExcelToDataTables(string excelPath) {
            if (!ExcelUtil.IsValidExcelFile(excelPath)) {
                return null;
            }
            Stream stream = new FileStream(excelPath, FileMode.Open);
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet dataSet = reader.AsDataSet();
            DataTableCollection dataTables = dataSet.Tables;
            return dataTables;
        }
    }
}
