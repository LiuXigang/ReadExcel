using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ReadExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            const string path = @"C:\Users\DELL\Desktop\ExcelDemo.xls";
            var str = "";
            //根据Excel生成C#类属性
            foreach (DataRow row in ExcelToTable(path).Rows)
            {
                var item1 = row.ItemArray[0];//参数   
                var item2 = row.ItemArray[1];//说明   
                var item3 = row.ItemArray[2];//类型   
                str += $"[DisplayName(\"{item2}\")]" + Environment.NewLine;
                str += $"public {ConvertType(item3.ToString())} {item1} ";
                str += "{ get; set;}" + Environment.NewLine;

            }
            Console.WriteLine(str);
            Console.ReadKey();
        }

        /// <summary>
        /// Excel导入成DataTble
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        private static DataTable ExcelToTable(string file)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                switch (fileExt)
                {
                    case ".xlsx":
                        workbook = new XSSFWorkbook(fs);
                        break;
                    case ".xls":
                        workbook = new HSSFWorkbook(fs);
                        break;
                    default:
                        workbook = null;
                        break;
                }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                default:
                    return "=" + cell.CellFormula;
            }
        }

        /// <summary>
        /// 装换Java类型为C#类型
        /// </summary>
        /// <param name="type">Java基本数据类型</param>
        /// <returns></returns>
        private static string ConvertType(string type)
        {
            string returnStr;
            switch (type)
            {
                case "String":
                    returnStr = "string"; break;
                case "Date":
                case "date":
                    returnStr = "DateTime"; break;
                case "Integer":
                case "integer":
                    returnStr = "int"; break;
                case "List":
                case "list":
                    returnStr = "List<>"; break;
                default:
                    returnStr = type; break;
            }
            return returnStr;
        }
    }
}
