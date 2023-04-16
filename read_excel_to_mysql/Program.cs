using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using MySql.Data.MySqlClient;
using System.Data;

namespace 学生成绩读取
{
    class Program
    {

        /// <summary>
        /// 读取Execl数据到DataTable(DataSet)中
        /// </summary>
        /// <param name="filePath">指定Execl文件路径</param>
        /// <param name="isFirstLineColumnName">设置第一行是否是列名</param>
        /// <returns>返回一个DataTable数据集</returns>
        public static DataSet ExcelToDataSet(string filePath, bool isFirstLineColumnName)
        { 
            DataSet dataSet = new DataSet();
            int startRow = 0;
            try
            {
                using (FileStream fs = File.OpenRead(filePath))
                {
                    IWorkbook workbook = null;
                    #region 兼容版本
                    // 如果是2007+的Excel版本
                    if (filePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // 如果是2003-的Excel版本
                    else if (filePath.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    #endregion

                    if (workbook != null)
                    {
                        //循环读取Excel的每个sheet，每个sheet页都转换为一个DataTable，并放在DataSet中
                        for (int p = 0; p < workbook.NumberOfSheets; p++)
                        {
                            ISheet sheet = workbook.GetSheetAt(p);
                            DataTable dataTable = new DataTable();
                            dataTable.TableName = sheet.SheetName;
                            if (sheet != null)
                            {
                                int rowCount = sheet.LastRowNum;//获取总行数
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(0);//获取第一行
                                    int cellCount = firstRow.LastCellNum;//获取总列数

                                    #region 为DataTabel添加列名
                                    if (isFirstLineColumnName)
                                    {
                                        startRow = 1;//如果第一行是列名，则从第二行开始读取
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)//获取列名
                                        {
                                            ICell cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    DataColumn column = new DataColumn(cell.StringCellValue);
                                                    dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)//如果纯数据，则默认以column1，column2形式设置列名
                                        {
                                            DataColumn column = new DataColumn("column" + (i + 1));
                                            dataTable.Columns.Add(column);
                                        }
                                    }
                                    #endregion

                                    # region 为DataSet添加数据行
                                    for (int i = startRow; i <= rowCount; ++i)
                                    {
                                        IRow row = sheet.GetRow(i);
                                        if (row == null) continue;//如果某行为空则跳转下一行

                                        DataRow dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            ICell cell = row.GetCell(j);
                                            if (cell == null)//如果单元格为空则以""空字符进行存储
                                            {
                                                dataRow[j] = "";
                                            }
                                            else
                                            {
                                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank://单元格设置为空，内部定义celltype。blank=3
                                                        dataRow[j] = "";
                                                        break;

                                                    case CellType.Numeric://单元格为数值形式

                                                        //获取判断是否为日期格式
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                        if (format == 14 || format == 22 || format == 31 || format == 57 || format == 58)
                                                            dataRow[j] = cell.DateCellValue;
                                                      //不为日期格式，则直接以数字存储
                                                        else
                                                            dataRow[j] = cell.NumericCellValue;
                                                        break;
                                                        //字符串形式则以字符串存储
                                                    case CellType.String:
                                                        dataRow[j] = cell.StringCellValue;
                                                        break;
                                                }
                                            }
                                        }
                                        dataTable.Rows.Add(dataRow);//为dataTable添加行
                                    }
                                    #endregion
                                }
                            }
                            dataSet.Tables.Add(dataTable);//为dataset添加datatable
                        }

                    }
                }
                return dataSet;
            }
            catch (Exception)
            {
               // var msg = ex.Message;
                return null;
            }
        }

        /// <summary>
        /// 将DataTable(DataSet)导出到Execl文档
        /// </summary>
        /// <param name="dataSet">传入一个DataSet</param>
        /// <param name="Outpath">导出路径（可以不加扩展名，不加默认为.xls）</param>
        /// <returns>返回一个Bool类型的值，表示是否导出成功</returns>
        /// True表示导出成功，Flase表示导出失败
        public static bool DataTableToExcel(DataSet dataSet, string Outpath)
        {
            bool result = false;
            try
            {
                if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0 || string.IsNullOrEmpty(Outpath))
                    throw new Exception("输入的DataSet或路径异常");
                int sheetIndex = 0;
                //根据输出路径的扩展名判断workbook的实例类型
                IWorkbook workbook = null;
                string pathExtensionName = Outpath.Trim().Substring(Outpath.Length - 5);
                if (pathExtensionName.Contains(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else if (pathExtensionName.Contains(".xls"))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    Outpath = Outpath.Trim() + ".xls";
                    workbook = new HSSFWorkbook();
                }
               // 将DataSet导出为Excel
                foreach (DataTable dt in dataSet.Tables)
                {
                    sheetIndex++;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(dt.TableName) ? ("sheet" + sheetIndex) : dt.TableName);//创建一个名称为Sheet0的表
                        int rowCount = dt.Rows.Count;//行数
                        int columnCount = dt.Columns.Count;//列数

                        //设置列头
                        IRow row = sheet.CreateRow(0);//excel第一行设为列头
                        for (int c = 0; c < columnCount; c++)
                        {
                            ICell cell = row.CreateCell(c);
                            cell.SetCellValue(dt.Columns[c].ColumnName);
                        }

                        //设置每行每列的单元格,
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < columnCount; j++)
                            {
                                ICell cell = row.CreateCell(j);//excel第二行开始写入数据
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                    }
                }
               // 向outPath输出数据
                using (FileStream fs = File.OpenWrite(Outpath))
                {
                    workbook.Write(fs);//向打开的这个xls文件中写入数据
                    result = true;
                }
   
                return result;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 将DataSet数据上传到Mysql数据库
        /// </summary>
        /// <param name="dataSet">获取的数据集</param>
        public static void DataTableToMysql(DataSet dataSet,string DataSheet)
            {
           //遍历每一个datatable
            foreach (DataTable dt in dataSet.Tables)
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    #region 创建数据表

                    //获取字段名
                    string[] row = new string[columnCount];
                    for (int c = 0; c < columnCount; c++)
                    {
                        row[c] = dt.Columns[c].ColumnName;
                    }
                  

                    string connStr = "server=localhost;port=3306;database=test;user=root;password=root;";

                    //生成创建数据表的sql'语句
                    string temp = "CREATE TABLE `"+DataSheet+"` (`";              
                    for (int i = 0; i < columnCount - 1; i++)
                    {
                        temp += (row[i] + "` VarChar(30),`");
                    }
                    temp += (row[columnCount - 1] + "`VarChar(30))");

                    //开始执行命令
                    using (MySqlConnection conn = new MySqlConnection(connStr))
                    {
                        conn.Open();
                        //防止第二次启动时再次新建数据表
                        try
                        {
                            //建表
                            using (MySqlCommand cmd = new MySqlCommand(temp, conn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                            Console.WriteLine("建表成功");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("建表失败，已存在");
                        }
                    }
                    #endregion

                    #region 获取数据内容，并上传数据库
                    //获取每行每列的值,并在获取一行数据后上传数据库
                    for (int i = 0; i < rowCount; i++)
                    {
                        //以字符串形式存储数据库
                        for (int j = 0; j < columnCount; j++)
                        {
                            row[j] = dt.Rows[i][j].ToString();
                        }
                        //创建sql语句
                        string sql = "insert into `"+DataSheet+"` value('";
                        for (int k = 0; k < columnCount - 1; k++)
                        {
                            sql += (row[k].ToString() + "','");
                        }
                        sql += (row[columnCount - 1].ToString() + "')");
                        using (MySqlConnection conn = new MySqlConnection(connStr))
                        {
                            conn.Open();
                            try
                            {
                                using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                Console.WriteLine("上传成功");
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("上传失败，已存在");
                            }
                        }
                    }
                    #endregion
                }

            }
        }
     
        static void Main(string[] args)
        {

            DataSet dataset = ExcelToDataSet(@"D:\Projects\Peojects_C#\学生成绩读取\学生成绩读取\学生成绩.xlsx", true);//使用绝对路径Excel导入              
            DataTableToMysql(dataset, "223");//将excel数据直接导入Mysql（自动建立数据表）
            if (DataTableToExcel(dataset, @"D:\Projects\Peojects_C#\学生成绩读取\学生成绩读取\1233.xlsx"))//绝对路径导出excel
            {
                Console.WriteLine("ok");
            }
        }
    }
} 
