using System;
using System.Data;
using System.IO;

using MySqlConnector;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TableTransfer
{
    public class TableTransfer
    {
        private const int DB_Addr = 0;
        private const int DB_Port = 1;
        private const int DB_User = 2;
        private const int DB_pwd = 3;
        private const int DB_Database = 4;
        private string tableName;
        private string cfgName;    //配置文件；
        private const string logName = "OP.log"; //访问日志；
        private string xlsxName; //Excel文件名；
        private string[] connString = new string[5]; //MySQL连接字符串；
        private DataTable dt;
        public TableTransfer(string CFGFile, string XLSXFile)
        {
            cfgName = CFGFile;
            using (StreamReader sr = new StreamReader(cfgName))
            {
                while (!sr.EndOfStream)
                {
                    string temp = sr.ReadLine();
                    switch (temp)
                    {
                        case "[Address]":
                            connString[DB_Addr] = "Addr=" + sr.ReadLine();
                            break;
                        case "[Port]":
                            connString[DB_Port] = "Port=" + sr.ReadLine();
                            break;
                        case "[User]":
                            connString[DB_User] = "User=" + sr.ReadLine();
                            break;
                        case "[Password]":
                            connString[DB_pwd] = "pwd=" + sr.ReadLine();
                            break;
                        case "[DatabaseName]":
                            connString[DB_Database] = "Database=" + sr.ReadLine().ToUpper();
                            break;
                        case "[TableName]":
                            tableName = sr.ReadLine().ToUpper();
                            break;
                        default:
                            break;
                    }
                }
            }
            xlsxName = XLSXFile;
            dt = new DataTable();
        }
        private static object GetValueType(ICell iCell)
        {   //获取
            if (iCell == null) return null;
            switch (iCell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return iCell.BooleanCellValue;
                case CellType.Numeric:
                    return iCell.NumericCellValue;
                case CellType.String:
                    return iCell.StringCellValue;
                case CellType.Formula:
                    return iCell.NumericCellValue;
                case CellType.Error:
                    return iCell.ErrorCellValue;
                default:
                    return iCell.CellFormula;
            }
        }
        private void readFromExcel()
        {
            using (StreamWriter sw = new StreamWriter(logName, true))
            {
                try
                {
                    using (FileStream fs = new FileStream(xlsxName, FileMode.Open, FileAccess.Read))
                    {
                        //避免产生错误；
                        fs.Position = 0;
                        IWorkbook wb = new XSSFWorkbook(fs);
                        //读取Excel的表数
                        for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
                        {
                            ISheet iSheet = wb.GetSheetAt(sheetNum);
                            //为DataTable添加表头：
                            IRow iRow = iSheet.GetRow(iSheet.FirstRowNum);
                            for (int cellNum = 0; cellNum < iRow.LastCellNum; cellNum++)
                            {
                                if (GetValueType(iRow.GetCell(cellNum)) == null)
                                    dt.Columns.Add(new DataColumn("Column" + cellNum.ToString()));
                                dt.Columns.Add(new DataColumn(GetValueType(iRow.GetCell(cellNum)).ToString()));
                            }
                            sw.WriteLine(iRow.LastCellNum.ToString() + " cells are readed from Excel.");
                            //为DataTable添加表内容：
                            for (int i = iSheet.FirstRowNum + 1; i <= iSheet.LastRowNum; i++)
                            {

                                iRow = iSheet.GetRow(i);
                                DataRow dr = dt.NewRow();
                                for (int j = 0; j < iRow.LastCellNum; j++)
                                    dr[j] = GetValueType(iRow.GetCell(j));
                                dt.Rows.Add(dr);
                            }
                            wb.Close();
                            sw.WriteLine(iSheet.LastRowNum.ToString() + " rows are readed from Excel.");
                        }
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine("readFromExcel(): " + e.ToString());
                }
            }
        }
        private void readFromMySQL()
        {
            //string MySqlCmd = "Select OLD_IP , NEW_IP , NEW_MASK , NEW_GW , NEW_DNS from " + tableName;
            string MySqlCmd = "Select * from " + tableName;
            using (StreamWriter sw = new StreamWriter(logName, true))
            {
                try
                {
                    using (MySqlDataAdapter dbAdapter = new MySqlDataAdapter(MySqlCmd, String.Join(";", connString)))
                    {
                        sw.WriteLine(dbAdapter.Fill(dt) + " rows are readed from MySQL.");
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine("readFromMySQL(): " + e.ToString());
                }
            }
        }
        private void writeToExcel()
        {
            using (FileStream fs = new FileStream(xlsxName, FileMode.OpenOrCreate, FileAccess.Write))
            {
                using (StreamWriter sw = new StreamWriter(logName, true))
                {
                    try
                    {
                        IWorkbook wb = new XSSFWorkbook();
                        ISheet iSheet = wb.CreateSheet(tableName);  //为Excel添加表头；
                        IRow iRow = iSheet.CreateRow(0);
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            iRow.CreateCell(i).SetCellValue(dt.Columns[i].ToString());
                        }
                        //为Excel添加内容；
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            iRow = iSheet.CreateRow(i + 1);   //NPOI和DataTable计算行数时相差1；
                            for (int j = 0; j < dt.Rows[i].ItemArray.Length; j++)
                            {
                                iRow.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());  //只能实现文本类型；
                            }
                        }
                        wb.Write(fs);
                        sw.WriteLine(iSheet.LastRowNum.ToString() + " rows are writed to Excel.");
                        wb.Close();
                    }
                    catch (Exception e)
                    {
                        sw.WriteLine("writeToExcel(): " + e.ToString());
                    }
                }
            }
        }
        private void writeToMySQL()
        {
            //Confirm MySql "SHOW GLOBAL VARIABLES LIKE 'local_infile'" is ON ; Otherwise "SET GLOBAL local_infile = 'ON'";
            using (MySqlConnection MySqlConn = new MySqlConnection(String.Join(";", connString) + ";AllowLoadLocalInfile=true"))
            {
                using (StreamWriter sw = new StreamWriter(logName, true))
                {
                    try
                    {
                        MySqlConn.Open();
                        MySqlBulkCopy MySqlBC = new MySqlBulkCopy(MySqlConn);
                        MySqlBC.DestinationTableName = tableName;
                        MySqlBulkCopyResult result = MySqlBC.WriteToServer(dt);
                        if (result.Warnings.Count != 0)
                        {
                            foreach (var w in result.Warnings)
                            {
                                sw.WriteLine(w.ToString());
                            }
                        }
                        sw.WriteLine(result.RowsInserted + " rows are written to MySQL.");
                    }
                    catch (Exception e)
                    {
                        sw.WriteLine("writeToMySQL(): " + e.ToString());
                    }
                }
            }
        }
        public void ExcelToMySQL()
        {
            readFromExcel();
            writeToMySQL();
        }
        public void MySQLToExcel()
        {
            readFromMySQL();
            writeToExcel();
        }
    }
}