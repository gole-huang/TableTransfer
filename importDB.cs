using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using MySqlConnector;

namespace importExceltoDB
{
    public class importDB
    {
        private const int DB_Addr = 0;
        private const int DB_Port = 1;
        private const int DB_User = 2;
        private const int DB_pwd = 3;
        private const int DB_Database = 4;
        private string cfgName;    //配置文件；
        private const string logName = "import.log"; //访问日志；
        private string xlsxName; //Excel文件名；
        private string[] connString = new string[5]; //MySQL连接字符串；
        private DataTable dt;

        public importDB(string CFGFile, string XLSXFile)
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
        /**/
        private DataTable readFromExcel()
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(xlsxName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook wb = new XSSFWorkbook(fs);
                ISheet iSheet = wb.GetSheetAt(0);
                //为DataTable添加表头：
                IRow iRow = iSheet.GetRow(iSheet.FirstRowNum);
                for (int i = 0; i < iRow.LastCellNum; i++)
                {
                    if (GetValueType(iRow.GetCell(i)) == null )
                        dt.Columns.Add(new DataColumn("Column"+i.ToString()));
                    dt.Columns.Add(new DataColumn(GetValueType(iRow.GetCell(i)).ToString()));
                }
                //为DataTable添加表内容：
                for (int i = iSheet.FirstRowNum + 1 ; i < iSheet.LastRowNum ; i++)
                {
                    iRow = iSheet.GetRow(i);
                    DataRow dr = dt.NewRow();
                    for (int j = 0 ; j < iRow.LastCellNum ; j++)
                        dr[j] = GetValueType(iRow.GetCell(j));
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        private DataTable readFromMySQL()
        {
            DataTable dt = new DataTable();
            string MySqlCmd = "Select OLD_IP , NEW_IP , NEW_MASK , NEW_GW , NEW_DNS from IP_RELATIONSHIP";
            using (MySqlDataAdapter dbAdapter = new MySqlDataAdapter(MySqlCmd, String.Join(';', connString)))
            {
                using (StreamWriter sw = new StreamWriter(logName))
                {
                    sw.WriteLine("From MySQL read " + dbAdapter.Fill(dt) + " lines.");
                }
            }
            return dt;
        }
        private void writeToExcel(DataTable dt)
        {
            using (FileStream fs = new FileStream(xlsxName, FileMode.Create, FileAccess.Write))
            {
            }
        }
        private async void writeToMySQL(DataTable dt)
        {
            using (MySqlConnection MySqlConn = new MySqlConnection(String.Join(';', connString) + ";AllowLoadLocalInfile=true"))
            {
                await MySqlConn.OpenAsync();
                MySqlBulkCopy MySqlBC = new MySqlBulkCopy(MySqlConn);
                MySqlBC.DestinationTableName = "IP_RELATIONSHIP";
                MySqlBulkCopyResult result = await MySqlBC.WriteToServerAsync(dt);
                using (StreamWriter sw = new StreamWriter(logName))
                {
                    if (result.Warnings.Count != 0)
                    {
                        foreach (var w in result.Warnings)
                            sw.WriteLine(w.ToString());
                    }
                    sw.WriteLine("To MySQL write " + result.RowsInserted + " lines.");
                }
            }
        }
    }
}
