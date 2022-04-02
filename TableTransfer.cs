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
        private string cfgName;    //配置文件；
        private const string logName = "OP.log"; //访问日志；
        private string xlsxName; //Excel文件名；
        private string[] connString = new string[5]; //MySQL连接字符串；
        //private DataTable dt;
        private DataSet ds;
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
                            connString[DB_Database] = "Database=" + sr.ReadLine();
                            break;
                        default:
                            break;
                    }
                }
            }
            xlsxName = XLSXFile;
            //dt = new DataTable();
            ds = new DataSet();
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
                        DataTable dt = new DataTable();
                        //ds.Tables.Add();
                        IWorkbook wb = new XSSFWorkbook(fs);
                        //读取Excel的表数
                        for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
                        {
                            DataTable dt=new DataTable();
                            ISheet iSheet = wb.GetSheetAt(sheetNum);
                            //为DataTable添加表头：
                            IRow iRow = iSheet.GetRow(iSheet.FirstRowNum);
                            for (int cellNum = 0; cellNum < iRow.LastCellNum; cellNum++)
                            {
                                if (GetValueType(iRow.GetCell(cellNum)) == null)
                                    dt.Columns.Add(new DataColumn("Column" + cellNum.ToString()));
                                dt.Columns.Add(new DataColumn(GetValueType(iRow.GetCell(cellNum)).ToString()));
                            }
                            sw.WriteLine(iRow.LastCellNum.ToString() + " cells had been readed.");
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
                            ds.Tables.Add(dt);  //往数据集中添加新表；
                            sw.WriteLine(iSheet.LastRowNum.ToString() + " rows had been readed.");
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
            string MySqlCmd = "Select OLD_IP , NEW_IP , NEW_MASK , NEW_GW , NEW_DNS from IP_RELATIONSHIP";
            using (StreamWriter sw = new StreamWriter(logName, true))
            {
                try
                {
                    using (MySqlDataAdapter dbAdapter = new MySqlDataAdapter(MySqlCmd, String.Join(";", connString)))
                    {
                        sw.WriteLine("From MySQL read " + dbAdapter.Fill(dt) + " lines.");
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
                        ISheet iSheet = wb.CreateSheet("IP_RELATIONSHIP");  //为Excel添加表头；
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
            using (MySqlConnection MySqlConn = new MySqlConnection(String.Join(";", connString) + ";AllowLoadLocalInfile=true"))
            {
                using (StreamWriter sw = new StreamWriter(logName, true))
                {
                    try
                    {
                        MySqlConn.Open();
                        MySqlBulkCopy MySqlBC = new MySqlBulkCopy(MySqlConn);
                        MySqlBC.DestinationTableName = "IP_RELATIONSHIP";
                        MySqlBulkCopyResult result = MySqlBC.WriteToServer(dt);
                        if (result.Warnings.Count != 0)
                        {
                            foreach (var w in result.Warnings)
                            {
                                sw.WriteLine(w.ToString());
                            }
                        }
                        sw.WriteLine(result.RowsInserted + " lines had been written to MySQL.");
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